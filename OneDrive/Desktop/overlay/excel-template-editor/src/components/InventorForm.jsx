import { inventorFields, inventorStartRow } from '../excel/mapping'
import { getCellByA1 } from '../excel/workbook'
import { useWorkbookStore } from '../state/useWorkbookStore'

function a1FromRowCol(row, colLetter) {
  return `${colLetter}${row}`
}

export default function InventorForm({ sheetName }) {
  const { workbook, setCell } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    setCell: state.setCell,
  }))

  const ws = workbook?.Sheets?.[sheetName]
  const inventors = Array.from({ length: 8 }, (_, index) => {
    const inventorIndex = index + 1
    const startRow = inventorStartRow(index)
    return { inventorIndex, startRow }
  })

  return (
    <section className="space-y-4">
      <div className="rounded-xl border border-slate-600 bg-gradient-to-r from-slate-900 to-slate-700 px-4 py-3 text-slate-100">
        <h2 className="font-['Sora'] text-lg font-semibold">Inventor information</h2>
        <p className="text-xs text-slate-300">Accordion form editing for inventor blocks.</p>
      </div>

      {inventors.map(({ inventorIndex, startRow }) => {
        const filled = inventorFields.filter((field) => {
          const row = startRow + field.offset
          const v = ws ? getCellByA1(ws, a1FromRowCol(row, 'B')) : ''
          return v != null && String(v).trim() !== ''
        }).length

        return (
          <details key={inventorIndex} className="rounded-2xl border border-slate-700 bg-slate-900/60 p-4" open={inventorIndex === 1}>
            <summary className="cursor-pointer font-['Sora'] text-base font-semibold text-slate-100">
              Inventor {inventorIndex}
              <span className="ml-2 text-xs font-medium text-slate-400">{filled}/{inventorFields.length} fields</span>
            </summary>
            <div className="mt-4 grid grid-cols-1 gap-4 md:grid-cols-2">
              {inventorFields.map((field) => {
                const row = startRow + field.offset
                const a1 = a1FromRowCol(row, 'B')
                const value = ws ? getCellByA1(ws, a1) : ''
                return (
                  <label key={`${inventorIndex}-${field.key}`} className="flex flex-col gap-1.5">
                    <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">{field.label}</span>
                    <input
                      type="text"
                      value={value ?? ''}
                      onChange={(event) => setCell(sheetName, a1, event.target.value)}
                      className="rounded-xl border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100 outline-none ring-sky-500 focus:ring"
                    />
                  </label>
                )
              })}
            </div>
            <div className="mt-4">
              <button
                type="button"
                onClick={() => {
                  inventorFields.forEach((field) => {
                    const row = startRow + field.offset
                    const a1 = a1FromRowCol(row, 'B')
                    setCell(sheetName, a1, '')
                  })
                }}
                className="rounded-xl bg-rose-600 px-3 py-2 text-sm font-semibold text-white hover:bg-rose-500"
              >
                Clear inventor {inventorIndex}
              </button>
            </div>
          </details>
        )
      })}
    </section>
  )
}
