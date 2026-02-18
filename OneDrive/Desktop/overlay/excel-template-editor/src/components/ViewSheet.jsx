import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { actionsSections, inventorFields, inventorStartRow, matterDetailsFields } from '../excel/mapping'
import { getCellByA1 } from '../excel/workbook'
import { useWorkbookStore } from '../state/useWorkbookStore'

function formatDisplay(value) {
  if (value == null || value === '') return '—'
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  return String(value)
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

function detectActionStages(ws) {
  const stages = []
  for (let col = 2; col <= 26; col += 1) {
    const c = getColumnLetter(col)
    let label = ''
    for (let row = 1; row <= 5; row += 1) {
      const raw = getCellByA1(ws, `${c}${row}`)
      const text = raw == null ? '' : String(raw).trim()
      if (text) {
        label = text
        break
      }
    }

    if (!label) {
      for (let row = 2; row <= 40; row += 1) {
        const raw = getCellByA1(ws, `${c}${row}`)
        const text = raw == null ? '' : String(raw).trim()
        if (text) {
          label = `Stage ${c}`
          break
        }
      }
    }

    if (!label) continue
    stages.push({ key: `${c}_${label}`, label, column: c })
  }

  return stages
}

function MatterViewer({ ws }) {
  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-600 bg-gradient-to-r from-slate-900 to-slate-700 px-4 py-3 text-slate-100">
        <h3 className="font-['Sora'] text-xl font-semibold">Matter details</h3>
        <p className="text-xs text-slate-300">Core file metadata pulled from the workbook.</p>
      </div>
      <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
        {matterDetailsFields.map((field) => (
          <div key={field.key} className="rounded-xl border border-slate-700 bg-slate-900/70 px-3 py-2 shadow-soft">
            <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">{field.label}</p>
            <p className="mt-1 rounded-lg bg-slate-800 px-2 py-1 text-sm font-medium text-slate-100">
              {formatDisplay(getCellByA1(ws, field.a1))}
            </p>
          </div>
        ))}
      </div>
    </div>
  )
}

function InventorViewer({ ws }) {
  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-600 bg-gradient-to-r from-slate-900 to-slate-700 px-4 py-3 text-slate-100">
        <h3 className="font-['Sora'] text-xl font-semibold">Inventor information</h3>
        <p className="text-xs text-slate-300">Scroll review of inventor profile blocks.</p>
      </div>
      {Array.from({ length: 8 }, (_, idx) => {
        const startRow = inventorStartRow(idx)
        const entries = inventorFields.map((field) => {
          const a1 = `B${startRow + field.offset}`
          return { label: field.label, value: getCellByA1(ws, a1) }
        })
        const hasAny = entries.some((e) => e.value != null && String(e.value).trim() !== '')
        if (!hasAny) return null

        return (
          <div key={idx} className="rounded-xl border border-slate-700 bg-slate-900/70 p-3 shadow-soft">
            <p className="mb-2 rounded-lg bg-slate-800 px-3 py-1.5 text-sm font-semibold text-slate-100">Inventor {idx + 1}</p>
            <div className="grid grid-cols-1 gap-2 md:grid-cols-2">
              {entries.map((entry) => (
                <div key={entry.label} className="rounded-lg bg-slate-800/80 px-3 py-2">
                  <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">{entry.label}</p>
                  <p className="mt-1 text-sm text-slate-100">{formatDisplay(entry.value)}</p>
                </div>
              ))}
            </div>
          </div>
        )
      })}
    </div>
  )
}

function ActionsViewer({ ws }) {
  const stages = useMemo(() => detectActionStages(ws), [ws])
  const [stageKey, setStageKey] = useState(stages[0]?.key || '')
  const selected = stages.find((s) => s.key === stageKey) || stages[0]

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div className="rounded-2xl border border-slate-600 bg-gradient-to-r from-slate-900 to-slate-700 px-4 py-3 text-slate-100">
          <h3 className="font-['Sora'] text-xl font-semibold">Actions overview</h3>
          <p className="text-xs text-slate-300">Stage-based read view with prompts separated from values.</p>
        </div>
        <label className="flex items-center gap-2 text-sm text-slate-300">
          Stage
          <select
            value={selected?.key || ''}
            onChange={(e) => setStageKey(e.target.value)}
            className="rounded-xl border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100"
          >
            {stages.map((s) => (
              <option key={s.key} value={s.key}>
                {s.label}
              </option>
            ))}
          </select>
        </label>
      </div>

      {actionsSections.map((section) => (
        <section key={section.key} className="rounded-xl border border-slate-700 bg-slate-900/70 p-3 shadow-soft">
          <h4 className="mb-3 rounded-lg border border-slate-700 bg-slate-800 px-3 py-2 text-sm font-semibold uppercase tracking-wide text-slate-200">
            {section.label}
          </h4>
          <div className="space-y-2">
            {Array.from({ length: section.endRow - section.startRow + 1 }, (_, i) => section.startRow + i).map((row) => {
              const prompt = getCellByA1(ws, `A${row}`)
              const value = selected ? getCellByA1(ws, `${selected.column}${row}`) : ''
              return (
                <div key={row} className="grid grid-cols-1 gap-2 rounded-lg border border-slate-700 bg-slate-800/80 p-3 md:grid-cols-[280px_1fr]">
                  <div className="rounded-md bg-slate-700 px-2 py-1">
                    <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-300">{formatDisplay(prompt)}</p>
                  </div>
                  <p className="whitespace-pre-wrap rounded-md bg-slate-900 px-2 py-1 text-sm text-slate-100">
                    {formatDisplay(value)}
                  </p>
                </div>
              )
            })}
          </div>
        </section>
      ))}
    </div>
  )
}

function GenericViewer({ ws }) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  return (
    <div className="overflow-auto rounded-xl border border-slate-700 bg-slate-900/70">
      <table className="min-w-full text-sm">
        <tbody>
          {data.map((row, ri) => (
            <tr key={ri} className="border-b border-slate-700">
              {row.map((cell, ci) => (
                <td key={ci} className="px-3 py-2 text-slate-200">
                  {formatDisplay(cell)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export default function ViewSheet({ sheetName }) {
  const workbook = useWorkbookStore((state) => state.workbook)
  const ws = workbook?.Sheets?.[sheetName]

  if (!ws) return <div className="text-sm text-slate-400">No sheet selected.</div>

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-sky-700/50 bg-sky-900/30 px-3 py-2 text-xs text-sky-100">
        View mode is optimized for reading and scroll-based review. Switch to FORM/GRID when editing.
      </div>

      {sheetName === 'Matter Details' && <MatterViewer ws={ws} />}
      {sheetName === 'Inventor Information - COMPLETE' && <InventorViewer ws={ws} />}
      {sheetName === 'Actions' && <ActionsViewer ws={ws} />}
      {!['Matter Details', 'Inventor Information - COMPLETE', 'Actions'].includes(sheetName) && <GenericViewer ws={ws} />}
    </div>
  )
}
