import { useMemo, useState } from 'react'
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
  const match = value.match(/^(High|Med|Low)\s+[—-]\s*(.*)$/)
  if (!match) return { rating: '', text: value }
  return { rating: match[1], text: match[2] || '' }
}

function composeRatingText(rating, text) {
  if (!rating) return text
  return `${rating} — ${text}`.trim()
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

export default function ActionsOverlay({ sheetName }) {
  const { workbook, setCell } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    setCell: state.setCell,
  }))
  const ws = workbook?.Sheets?.[sheetName]

  const stageOptions = useMemo(() => detectStageOptions(ws), [ws])

  const [stageKey, setStageKey] = useState(stageOptions[0]?.key ?? DEFAULT_STAGES[0].key)
  const [showContinuationForAll, setShowContinuationForAll] = useState(false)
  const [copyFromStageKey, setCopyFromStageKey] = useState(stageOptions[0]?.key ?? DEFAULT_STAGES[0].key)

  const selectedStage = stageOptions.find((s) => s.key === stageKey) || stageOptions[0]
  const selectedCopyFromStage = stageOptions.find((s) => s.key === copyFromStageKey) || stageOptions[0]
  const stageColumn = selectedStage?.column ?? 'B'

  const visibleSections = actionsSections.filter((section) => {
    if (section.key !== 'continuation') return true
    return showContinuationForAll
  })

  const updateRowValue = (row, nextValue, rating = null) => {
    const a1 = `${stageColumn}${row}`
    if (TRIAD_SET.has(row)) {
      const current = ws ? getCellByA1(ws, a1) : ''
      const parsed = parseRatingText(current)
      const nextRating = rating === null ? parsed.rating : rating
      setCell(sheetName, a1, composeRatingText(nextRating, nextValue))
      return
    }

    if (row === DATE_ROW) {
      setCell(sheetName, a1, nextValue === '' ? '' : nextValue)
      return
    }

    setCell(sheetName, a1, nextValue)
  }

  const copyMainStrategyFromStage = () => {
    if (!ws || !selectedCopyFromStage || !selectedStage) return
    if (selectedCopyFromStage.column === selectedStage.column) return

    for (let row = 6; row <= 17; row += 1) {
      const source = getCellByA1(ws, `${selectedCopyFromStage.column}${row}`)
      setCell(sheetName, `${stageColumn}${row}`, source ?? '')
    }
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
                const baseValue = parsed ? parsed.text : rawValue == null ? '' : String(rawValue)

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
                            onClick={() => updateRowValue(row, parsed?.text ?? '', pill)}
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
                          onClick={() => updateRowValue(row, parsed?.text ?? '', '')}
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
                      <textarea
                        value={baseValue}
                        onChange={(event) => updateRowValue(row, event.target.value)}
                        rows={6}
                        className="w-full rounded-xl border border-slate-600 bg-slate-900 px-3 py-2 text-slate-100"
                      />
                    ) : kind === 'date' ? (
                      <input
                        type="date"
                        value={formatDateValue(baseValue)}
                        onChange={(event) => updateRowValue(row, event.target.value)}
                        className="w-full rounded-xl border border-slate-600 bg-slate-900 px-3 py-2 text-slate-100"
                      />
                    ) : (
                      <input
                        type="text"
                        value={baseValue}
                        onChange={(event) => updateRowValue(row, event.target.value)}
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
