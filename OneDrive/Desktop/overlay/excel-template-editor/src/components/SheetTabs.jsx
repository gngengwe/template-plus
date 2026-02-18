import { useWorkbookStore } from '../state/useWorkbookStore'

export default function SheetTabs({ sheetNames }) {
  const { activeSheetName, setActiveSheet } = useWorkbookStore((state) => ({
    activeSheetName: state.activeSheetName,
    setActiveSheet: state.setActiveSheet,
  }))

  if (!sheetNames?.length) return <div className="text-sm text-slate-400">No workbook loaded</div>

  return (
    <div className="flex flex-wrap gap-2">
      {sheetNames.map((name) => (
        <button
          type="button"
          key={name}
          onClick={() => setActiveSheet(name)}
          className={`rounded-xl border px-3 py-1.5 text-sm font-medium transition ${
            activeSheetName === name
              ? 'border-sky-400/60 bg-sky-500/20 text-sky-100'
              : 'border-slate-700 bg-slate-800 text-slate-300 hover:border-slate-500 hover:text-slate-100'
          }`}
        >
          {name}
        </button>
      ))}
    </div>
  )
}
