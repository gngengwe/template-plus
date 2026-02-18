import { useMemo } from 'react'
import { HotTable } from '@handsontable/react'
import { registerAllModules } from 'handsontable/registry'
import * as XLSX from 'xlsx'
import { useWorkbookStore } from '../state/useWorkbookStore'

registerAllModules()

export default function GridSheet({ sheetName }) {
  const { workbook, setCell } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    setCell: state.setCell,
  }))

  const ws = workbook?.Sheets?.[sheetName]

  const data = useMemo(() => {
    if (!ws) return [[]]
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  }, [ws])

  if (!ws) return <div className="text-sm text-slate-400">No sheet selected.</div>

  return (
    <div className="hot-container overflow-hidden rounded-2xl border border-slate-700 bg-white shadow-soft">
      <HotTable
        data={data}
        rowHeaders
        colHeaders
        licenseKey="non-commercial-and-evaluation"
        height="65vh"
        stretchH="all"
        afterChange={(changes, source) => {
          if (!changes || source === 'loadData') return
          changes.forEach(([row, col, oldValue, newValue]) => {
            if (oldValue === newValue) return
            const a1 = XLSX.utils.encode_cell({ r: row, c: col })
            setCell(sheetName, a1, newValue ?? '')
          })
        }}
      />
    </div>
  )
}
