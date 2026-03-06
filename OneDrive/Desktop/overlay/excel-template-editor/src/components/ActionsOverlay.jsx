import { useMemo, useRef, useState } from 'react'
import { actionsSections, triadRows } from '../excel/mapping'
import { getCellByA1 } from '../excel/workbook'
import { useWorkbookStore } from '../state/useWorkbookStore'

const TEXTAREA_ROWS = new Set([6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 28])
const SHORT_ROWS = new Set([2, 4])
const DATE_ROW = 3
const TRIAD_SET = new Set([...triadRows.main, ...triadRows.continuation])
const RATINGS = ['High', 'Med', 'Low']
const DEFAULT_STAGES = [
  { key: 'ENGAGEMENT_B', label: 'Engagement', column: 'B' },
  { key: 'DISCLOSURE_PRE_C', label: 'Disclosure: Pre-Search', column: 'C' },
  { key: 'DISCLOSURE_POST_D', label: 'Disclosure: Post Search', column: 'D' },
  { key: 'RESTRICTION_E', label: 'Restriction', column: 'E' },
  { key: 'OFFICE_ACTION_F', label: 'Office Action / Appeal Decision', column: 'F' },
  { key: 'ADVISORY_AFCP_G', label: "Advisory/AFCP/Pre-Appeal/Examiner's Answer", column: 'G' },
  { key: 'NOA_H', label: 'Notice of Allowance', column: 'H' },
]

function parseRatingText(value) {
  if (typeof value !== 'string') return { rating: '', text: value == null ? '' : String(value) }
  const match = value.match(/^(High|Med|Low)\s+[�-]\s*(.*)$/)
  if (!match) return { rating: '', text: value }
  return { rating: match[1], text: match[2] || '' }
}

function composeRatingText(rating, text) {
  if (!rating) return text
  return `${rating} � ${text}`.trim()
}

function inputKindForRow(row) {
  if (row === DATE_ROW) return 'date'
  if (SHORT_ROWS.has(row)) return 'short'
  if (TEXTAREA_ROWS.has(row)) return 'long'
  return 'short'
}

function formatDateValue(value) {
  if (!value) return ''
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  if (typeof value === 'number' && Number.isFinite(value)) {
    const utcDays = Math.floor(value - 25569)
    const utcValue = utcDays * 86400
    const dateInfo = new Date(utcValue * 1000)
    if (!Number.isNaN(dateInfo.getTime())) return dateInfo.toISOString().slice(0, 10)
  }
  if (typeof value === 'string') return value.slice(0, 10)
  return ''
}

function sanitizeKey(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
}

function detectStageOptions(ws) {
  if (!ws) return DEFAULT_STAGES

  const options = []
  for (let col = 2; col <= 26; col += 1) {
    const columnLetter = getColumnLetter(col)
    const header = findHeaderLabel(ws, columnLetter)
    const hasStageData = columnHasStageData(ws, columnLetter)
    if (!header && !hasStageData) continue

    options.push({
      key: `${sanitizeKey(header || `Stage ${columnLetter}`) || 'STAGE'}_${columnLetter}`,
      label: header || `Stage ${columnLetter}`,
      column: columnLetter,
    })
  }

  if (options.length === 0) return DEFAULT_STAGES

  const counts = {}
  const seen = {}
  options.forEach((opt) => {
    counts[opt.label] = (counts[opt.label] || 0) + 1
  })

  return options.map((opt) => {
    seen[opt.label] = (seen[opt.label] || 0) + 1
    if ((counts[opt.label] || 0) <= 1) return opt
    return {
      ...opt,
      label: `${opt.label} (${seen[opt.label]})`,
    }
  })
}

function findHeaderLabel(ws, columnLetter) {
  for (let row = 1; row <= 5; row += 1) {
    const raw = getCellByA1(ws, `${columnLetter}${row}`)
    const text = raw == null ? '' : String(raw).trim()
    if (text) return text
  }
  return ''
}

function columnHasStageData(ws, columnLetter) {
  for (let row = 2; row <= 40; row += 1) {
    const raw = getCellByA1(ws, `${columnLetter}${row}`)
    const text = raw == null ? '' : String(raw).trim()
    if (text) return true
  }
  return false
}

function getColumnLetter(columnNumber) {
  let n = columnNumber
  let out = ''
  while (n > 0) {
    const mod = (n - 1) % 26
    out = String.fromCharCode(65 + mod) + out
    n = Math.floor((n - mod) / 26)
  }
  return out
}

function isBroadestClaimField(label, row) {
  const text = String(label || '').toLowerCase()
  if ((text.includes('broadest') && text.includes('claim')) || /broadest.*claim|claim.*broadest/i.test(text)) {
    return true
  }
  // Template fallback: the broadest-claim prompt is typically row 6 in Actions.
  return Number(row) === 6
}

function parseUsptoMarkup(value) {
  const source = String(value || '')
  const tokenRegex = /(\*\*[^*]+\*\*|__[^_]+__|~~[^~]+~~)/g
  const runs = []
  let plain = ''
  let cursor = 0
  let match = tokenRegex.exec(source)

  while (match) {
    if (match.index > cursor) {
      const text = source.slice(cursor, match.index)
      runs.push({ text, bold: false, underline: false, strike: false })
      plain += text
    }

    const token = match[0]
    if (token.startsWith('**') && token.endsWith('**')) {
      const text = token.slice(2, -2)
      runs.push({ text, bold: true, underline: false, strike: false })
      plain += text
    } else if (token.startsWith('__') && token.endsWith('__')) {
      const text = token.slice(2, -2)
      runs.push({ text, bold: false, underline: true, strike: false })
      plain += text
    } else if (token.startsWith('~~') && token.endsWith('~~')) {
      const text = token.slice(2, -2)
      runs.push({ text, bold: false, underline: false, strike: true })
      plain += text
    }

    cursor = match.index + token.length
    match = tokenRegex.exec(source)
  }

  if (cursor < source.length) {
    const text = source.slice(cursor)
    runs.push({ text, bold: false, underline: false, strike: false })
    plain += text
  }

  return {
    plainText: plain,
    runs: runs.filter((run) => run.text.length > 0),
  }
}

export default function ActionsOverlay({ sheetName }) {
  const { workbook, setCell } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    setCell: state.setCell,
  }))
  const ws = workbook?.Sheets?.[sheetName]

  const stageOptions = useMemo(() => detectStageOptions(ws), [ws])

  const [stageKey, setStageKey] = useState(stageOptions[0]?.key ?? DEFAULT_STAGES[0].key)
  const [showContinuationForAll, setShowContinuationForAll] = useState(false)
  const [showUsptoToolsForAllLongFields, setShowUsptoToolsForAllLongFields] = useState(false)
  const [copyFromStageKey, setCopyFromStageKey] = useState(stageOptions[0]?.key ?? DEFAULT_STAGES[0].key)
  const [richDraftByA1, setRichDraftByA1] = useState({})
  const editorRefs = useRef({})

  const selectedStage = stageOptions.find((s) => s.key === stageKey) || stageOptions[0]
  const selectedCopyFromStage = stageOptions.find((s) => s.key === copyFromStageKey) || stageOptions[0]
  const stageColumn = selectedStage?.column ?? 'B'

  const visibleSections = actionsSections.filter((section) => {
    if (section.key !== 'continuation') return true
    return showContinuationForAll
  })

  const updateRowValue = (row, nextValue, rating = null, rowLabel = '', useUsptoMarkup = false) => {
    const a1 = `${stageColumn}${row}`
    if (TRIAD_SET.has(row)) {
      const current = ws ? getCellByA1(ws, a1) : ''
      const parsed = parseRatingText(current)
      const nextRating = rating === null ? parsed.rating : rating
      setCell(sheetName, a1, composeRatingText(nextRating, nextValue))
      return
    }

    if (row === DATE_ROW) {
      if (nextValue === '') {
        setCell(sheetName, a1, '', { styleKey: 'actionsDate', forceType: 'd' })
        return
      }

      const dateValue = new Date(`${nextValue}T00:00:00`)
      if (!Number.isNaN(dateValue.getTime())) {
        setCell(sheetName, a1, dateValue, { styleKey: 'actionsDate', forceType: 'd' })
        return
      }

      setCell(sheetName, a1, nextValue, { styleKey: 'actionsDate', forceType: 'd' })
      return
    }

    if (isBroadestClaimField(rowLabel, row) || useUsptoMarkup) {
      const parsedMarkup = parseUsptoMarkup(nextValue)
      setCell(sheetName, a1, parsedMarkup.plainText, { richTextRuns: parsedMarkup.runs })
      return
    }

    setCell(sheetName, a1, nextValue, { richTextRuns: null })
  }

  const copyMainStrategyFromStage = () => {
    if (!ws || !selectedCopyFromStage || !selectedStage) return
    if (selectedCopyFromStage.column === selectedStage.column) return

    for (let row = 6; row <= 17; row += 1) {
      const source = getCellByA1(ws, `${selectedCopyFromStage.column}${row}`)
      setCell(sheetName, `${stageColumn}${row}`, source ?? '')
    }
  }

  const getBroadestDraftValue = (a1, rawValue) => {
    if (Object.prototype.hasOwnProperty.call(richDraftByA1, a1)) return richDraftByA1[a1]
    return rawValue == null ? '' : String(rawValue)
  }

  const commitBroadestDraft = (row, rowLabel, a1, draftValue, useUsptoMarkup = false) => {
    setRichDraftByA1((prev) => ({ ...prev, [a1]: draftValue }))
    updateRowValue(row, draftValue, null, rowLabel, useUsptoMarkup)
  }

  const applyWrapperToBroadest = (row, rowLabel, a1, openMark, closeMark = openMark) => {
    const el = editorRefs.current[a1]
    const current = getBroadestDraftValue(a1, ws ? getCellByA1(ws, a1) : '')
    const start = el?.selectionStart ?? current.length
    const end = el?.selectionEnd ?? current.length
    const selected = current.slice(start, end)
    const wrapped = `${openMark}${selected}${closeMark}`
    const nextValue = `${current.slice(0, start)}${wrapped}${current.slice(end)}`
    commitBroadestDraft(row, rowLabel, a1, nextValue, true)

    const nextCursorStart = start + openMark.length
    const nextCursorEnd = nextCursorStart + selected.length
    requestAnimationFrame(() => {
      const nextEl = editorRefs.current[a1]
      if (!nextEl) return
      nextEl.focus()
      nextEl.setSelectionRange(nextCursorStart, nextCursorEnd)
    })
  }

  const insertTokenInBroadest = (row, rowLabel, a1, token) => {
    const el = editorRefs.current[a1]
    const current = getBroadestDraftValue(a1, ws ? getCellByA1(ws, a1) : '')
    const start = el?.selectionStart ?? current.length
    const end = el?.selectionEnd ?? current.length
    const nextValue = `${current.slice(0, start)}${token}${current.slice(end)}`
    commitBroadestDraft(row, rowLabel, a1, nextValue, true)

    const nextCursor = start + token.length
    requestAnimationFrame(() => {
      const nextEl = editorRefs.current[a1]
      if (!nextEl) return
      nextEl.focus()
      nextEl.setSelectionRange(nextCursor, nextCursor)
    })
  }

  return (
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-[150px_1fr]">
      <aside className="rounded-2xl border border-slate-700 bg-slate-900/60 p-3">
        <p className="mb-2 font-['Sora'] text-sm font-semibold text-slate-100">Jump to section</p>
        <div className="flex flex-col gap-2">
          {visibleSections.map((section) => (
            <a
              key={section.key}
              href={`#section-${section.key}`}
              className="rounded-xl border border-slate-700 bg-slate-800 px-3 py-2 text-sm text-slate-200 transition hover:border-slate-500"
            >
              {section.label}
            </a>
          ))}
        </div>
      </aside>

      <section className="space-y-4 rounded-2xl border border-slate-700 bg-slate-900/60 p-4">
        <div className="rounded-xl border border-slate-600 bg-gradient-to-r from-slate-900 to-slate-700 px-4 py-3 text-slate-100">
          <h2 className="font-['Sora'] text-lg font-semibold">Actions editor</h2>
          <p className="text-xs text-slate-300">Prompt and response editing with stage-aware controls and consistent field formatting.</p>
        </div>

        <div className="rounded-xl border border-sky-700/50 bg-sky-900/30 px-3 py-2 text-xs text-sky-100">
          Edited values save with Excel alignment <strong>General</strong> (horizontal), <strong>Top</strong> (vertical), and <strong>Wrap Text</strong> enabled.
        </div>

        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
          <label className="flex flex-col gap-1">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">Stage</span>
            <select
              value={selectedStage?.key || ''}
              onChange={(event) => setStageKey(event.target.value)}
              className="rounded-xl border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100"
            >
              {stageOptions.map((item) => (
                <option key={item.key} value={item.key}>
                  {item.label}
                </option>
              ))}
            </select>
          </label>

          <label className="flex flex-col gap-1">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">Copy strategy from</span>
            <select
              value={selectedCopyFromStage?.key || ''}
              onChange={(event) => setCopyFromStageKey(event.target.value)}
              className="rounded-xl border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100"
            >
              {stageOptions.map((item) => (
                <option key={item.key} value={item.key}>
                  {item.label}
                </option>
              ))}
            </select>
          </label>

          <div className="flex items-end">
            <button
              type="button"
              onClick={copyMainStrategyFromStage}
              className="w-full rounded-xl bg-blue-600 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-500"
            >
              Copy strategy
            </button>
          </div>
        </div>

        <label className="flex items-center gap-2 text-sm text-slate-300">
          <input
            type="checkbox"
            checked={showContinuationForAll}
            onChange={(event) => setShowContinuationForAll(event.target.checked)}
          />
          Show continuation proposal for this stage
        </label>

        <label className="flex items-center gap-2 text-sm text-slate-300">
          <input
            type="checkbox"
            checked={showUsptoToolsForAllLongFields}
            onChange={(event) => setShowUsptoToolsForAllLongFields(event.target.checked)}
          />
          Show USPTO tools on all long fields
        </label>

        {visibleSections.map((section) => (
          <div id={`section-${section.key}`} key={section.key} className="rounded-xl border border-slate-700 bg-slate-900/70 p-3 shadow-soft">
            <h3 className="mb-3 rounded-lg border border-slate-700 bg-slate-800 px-3 py-2 font-['Sora'] text-sm font-semibold uppercase tracking-wide text-slate-200">
              {section.label}
            </h3>

            <div className="space-y-3">
              {Array.from({ length: section.endRow - section.startRow + 1 }, (_, idx) => section.startRow + idx).map((row) => {
                const label = ws ? getCellByA1(ws, `A${row}`) : ''
                const a1 = `${stageColumn}${row}`
                const rawValue = ws ? getCellByA1(ws, a1) : ''
                const parsed = TRIAD_SET.has(row) ? parseRatingText(rawValue) : null
                const kind = inputKindForRow(row)
                const isBroadestClaim = isBroadestClaimField(label, row)
                const showUsptoToolbar = kind === 'long' && (isBroadestClaim || showUsptoToolsForAllLongFields)
                const baseValue = parsed
                  ? parsed.text
                  : showUsptoToolbar
                    ? getBroadestDraftValue(a1, rawValue)
                    : rawValue == null
                      ? ''
                      : String(rawValue)

                return (
                  <div key={row} className="rounded-lg border border-slate-700 bg-slate-800/70 p-3">
                    <div className="mb-2 rounded-md bg-slate-700/70 px-2 py-1">
                      <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-300">{label || `Row ${row}`}</p>
                    </div>

                    {TRIAD_SET.has(row) && (
                      <div className="mb-2 flex flex-wrap gap-2">
                        {RATINGS.map((pill) => (
                          <button
                            key={pill}
                            type="button"
                            onClick={() => updateRowValue(row, parsed?.text ?? '', pill, label)}
                            className={`rounded-full px-3 py-1 text-xs font-medium ${
                              parsed?.rating === pill
                                ? 'bg-sky-600 text-white'
                                : 'bg-slate-700 text-slate-200 ring-1 ring-slate-600 hover:bg-slate-600'
                            }`}
                          >
                            {pill}
                          </button>
                        ))}
                        <button
                          type="button"
                          onClick={() => updateRowValue(row, parsed?.text ?? '', '', label)}
                          className={`rounded-full px-3 py-1 text-xs font-medium ${
                            !parsed?.rating
                              ? 'bg-sky-600 text-white'
                              : 'bg-slate-700 text-slate-200 ring-1 ring-slate-600 hover:bg-slate-600'
                          }`}
                        >
                          None
                        </button>
                      </div>
                    )}

                    {kind === 'long' ? (
                      <>
                        {showUsptoToolbar && (
                          <div className="mb-2 flex flex-wrap gap-2">
                            <button
                              type="button"
                              onClick={() => applyWrapperToBroadest(row, label, a1, '__', '__')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              Underline
                            </button>
                            <button
                              type="button"
                              onClick={() => applyWrapperToBroadest(row, label, a1, '~~', '~~')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              Strikethrough
                            </button>
                            <button
                              type="button"
                              onClick={() => applyWrapperToBroadest(row, label, a1, '**', '**')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              Bold
                            </button>
                            <button
                              type="button"
                              onClick={() => applyWrapperToBroadest(row, label, a1, '[', ']')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              Brackets
                            </button>
                            <button
                              type="button"
                              onClick={() => insertTokenInBroadest(row, label, a1, '[AMENDED] ')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              [AMENDED]
                            </button>
                            <button
                              type="button"
                              onClick={() => insertTokenInBroadest(row, label, a1, '[CANCELED] ')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              [CANCELED]
                            </button>
                            <button
                              type="button"
                              onClick={() => insertTokenInBroadest(row, label, a1, '[NEW] ')}
                              className="rounded-md bg-slate-700 px-2 py-1 text-xs text-slate-100 hover:bg-slate-600"
                            >
                              [NEW]
                            </button>
                          </div>
                        )}

                        <textarea
                          value={baseValue}
                          onChange={(event) => {
                            if (showUsptoToolbar) {
                              commitBroadestDraft(row, label, a1, event.target.value, true)
                              return
                            }
                            updateRowValue(row, event.target.value, null, label)
                          }}
                          rows={6}
                          ref={(el) => {
                            if (showUsptoToolbar) editorRefs.current[a1] = el
                          }}
                          className="w-full rounded-xl border border-slate-600 bg-slate-900 px-3 py-2 text-slate-100"
                        />
                        {showUsptoToolbar && (
                          <p className="mt-2 text-[11px] text-slate-400">
                            USPTO markup: <code>__underline__</code>, <code>~~strikethrough~~</code>, <code>**bold**</code>, and <code>[brackets]</code> for deleted matter.
                          </p>
                        )}
                      </>
                    ) : kind === 'date' ? (
                      <input
                        type="date"
                        value={formatDateValue(baseValue)}
                        onChange={(event) => updateRowValue(row, event.target.value, null, label)}
                        className="w-full rounded-xl border border-slate-600 bg-slate-900 px-3 py-2 text-slate-100"
                      />
                    ) : (
                      <input
                        type="text"
                        value={baseValue}
                        onChange={(event) => updateRowValue(row, event.target.value, null, label)}
                        className="w-full rounded-xl border border-slate-600 bg-slate-900 px-3 py-2 text-slate-100"
                      />
                    )}
                  </div>
                )
              })}
            </div>
          </div>
        ))}
      </section>
    </div>
  )
}
