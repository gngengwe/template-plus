import { matterDetailsFields } from '../excel/mapping'
import { getCellByA1 } from '../excel/workbook'
import { useWorkbookStore } from '../state/useWorkbookStore'

function formatDateForInput(value) {
  if (!value) return ''
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  if (typeof value === 'string') return value.slice(0, 10)
  return ''
}

export default function MatterDetailsForm({ sheetName }) {
  const { workbook, setCell } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    setCell: state.setCell,
  }))

  const ws = workbook?.Sheets?.[sheetName]

  return (
    <section className="space-y-4 rounded-2xl border border-slate-700 bg-slate-900/60 p-5">
      <div className="rounded-xl border border-slate-600 bg-gradient-to-r from-slate-900 to-slate-700 px-4 py-3 text-slate-100">
        <h2 className="font-['Sora'] text-lg font-semibold">Matter details</h2>
        <p className="text-xs text-slate-300">Structured form editing for core matter metadata.</p>
      </div>

      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
        {matterDetailsFields.map((field) => {
          const rawValue = ws ? getCellByA1(ws, field.a1) : ''
          const value = field.type === 'date' ? formatDateForInput(rawValue) : rawValue ?? ''
          const isMissingRequired = Boolean(field.required && (!value || String(value).trim() === ''))

          return (
            <label key={field.key} className="flex flex-col gap-1.5">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">
                {field.label}
                {field.required ? ' *' : ''}
              </span>
              <input
                type={field.type === 'date' ? 'date' : field.type === 'number' ? 'number' : 'text'}
                value={value}
                onChange={(event) => {
                  const next = event.target.value
                  if (field.type === 'number') return setCell(sheetName, field.a1, next === '' ? '' : Number(next))
                  if (field.type === 'date') return setCell(sheetName, field.a1, next === '' ? '' : next)
                  return setCell(sheetName, field.a1, next)
                }}
                className={`rounded-xl border bg-slate-800 px-3 py-2 text-slate-100 outline-none ring-sky-500 focus:ring ${
                  isMissingRequired ? 'border-amber-400/80' : 'border-slate-600'
                }`}
              />
            </label>
          )
        })}
      </div>
    </section>
  )
}
