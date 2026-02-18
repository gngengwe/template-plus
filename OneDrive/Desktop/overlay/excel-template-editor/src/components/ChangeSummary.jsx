import { useMemo } from 'react'
import { useWorkbookStore } from '../state/useWorkbookStore'

function formatValue(value) {
  if (value === '' || value === null || value === undefined) return '(blank)'
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  const asString = String(value)
  return asString.length > 44 ? `${asString.slice(0, 44)}...` : asString
}

export default function ChangeSummary() {
  const { editsBySheet, dirty, activeSheetName, setActiveSheet } = useWorkbookStore((state) => ({
    editsBySheet: state.editsBySheet,
    dirty: state.dirty,
    activeSheetName: state.activeSheetName,
    setActiveSheet: state.setActiveSheet,
  }))

  const flattened = useMemo(() => {
    const rows = []
    Object.entries(editsBySheet || {}).forEach(([sheetName, byA1]) => {
      Object.entries(byA1 || {}).forEach(([a1, edit]) => {
        rows.push({ sheetName, a1, value: edit?.v })
      })
    })
    return rows
  }, [editsBySheet])

  const total = flattened.length

  return (
    <aside className="rounded-2xl border border-slate-700 bg-slate-900/70 p-4 shadow-soft">
      <div className="mb-3 flex items-center justify-between gap-2">
        <h3 className="font-['Sora'] text-base font-semibold text-slate-100">Change summary</h3>
        <span className={`rounded-full px-2 py-0.5 text-xs font-semibold ${dirty ? 'bg-amber-500/20 text-amber-200' : 'bg-emerald-500/20 text-emerald-200'}`}>
          {total}
        </span>
      </div>

      {!total ? (
        <p className="text-sm text-slate-400">No pending edits.</p>
      ) : (
        <div className="space-y-2">
          {flattened.slice(0, 16).map((item) => (
            <div key={`${item.sheetName}-${item.a1}`} className="rounded-xl border border-slate-700 bg-slate-800/70 px-3 py-2">
              <div className="mb-1 flex items-center justify-between gap-2">
                <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">{item.a1}</span>
                <button
                  type="button"
                  onClick={() => setActiveSheet(item.sheetName)}
                  className={`rounded-lg px-2 py-0.5 text-xs font-medium ${
                    activeSheetName === item.sheetName
                      ? 'bg-sky-500/20 text-sky-200'
                      : 'bg-slate-700 text-slate-200 ring-1 ring-slate-600 hover:bg-slate-600'
                  }`}
                >
                  {activeSheetName === item.sheetName ? 'Current' : 'Go'}
                </button>
              </div>
              <p className="text-xs font-medium text-slate-200">{item.sheetName}</p>
              <p className="mt-0.5 text-xs text-slate-400">{formatValue(item.value)}</p>
            </div>
          ))}

          {total > 16 && <p className="text-xs text-slate-400">Showing first 16 changes. Save to clear all pending edits.</p>}
        </div>
      )}
    </aside>
  )
}
