const MODES = ['AUTO', 'VIEW', 'FORM', 'GRID']

export default function ModeToggle({ mode, onModeChange }) {
  return (
    <div className="inline-flex rounded-2xl border border-slate-700 bg-slate-900 p-1 shadow-soft">
      {MODES.map((item) => (
        <button
          key={item}
          type="button"
          onClick={() => onModeChange(item)}
          className={`rounded-xl px-3 py-1.5 text-sm font-medium transition ${
            mode === item
              ? 'bg-sky-600 text-white'
              : 'text-slate-300 hover:bg-slate-800 hover:text-slate-100'
          }`}
        >
          {item}
        </button>
      ))}
    </div>
  )
}
