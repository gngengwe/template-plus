import { useMemo, useState } from 'react'
import { useWorkbookStore } from '../state/useWorkbookStore'

function isFsAccessSupported() {
  return typeof window !== 'undefined' && 'showOpenFilePicker' in window
}

function formatSavedTime(iso) {
  if (!iso) return ''
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
}

export default function Toolbar({ viewOnly = false }) {
  const {
    workbook,
    fileHandle,
    fileName,
    dirty,
    lastSavedAt,
    openFile,
    openFileWithMode,
    save,
    saveAs,
    downloadCopy,
  } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    fileHandle: state.fileHandle,
    fileName: state.fileName,
    dirty: state.dirty,
    lastSavedAt: state.lastSavedAt,
    openFile: state.openFile,
    openFileWithMode: state.openFileWithMode,
    save: state.save,
    saveAs: state.saveAs,
    downloadCopy: state.downloadCopy,
  }))

  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const supportsFs = useMemo(() => isFsAccessSupported(), [])
  const hasWorkbook = Boolean(workbook)
  const canSameFileSave = supportsFs && hasWorkbook && Boolean(fileHandle)

  const runAction = async (action) => {
    setError('')
    setBusy(true)
    try {
      await action()
    } catch (err) {
      setError(err?.message || 'Action failed.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <header className="rounded-3xl border border-slate-700/80 bg-slate-900/85 p-5 shadow-glow backdrop-blur">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="font-['Sora'] text-2xl font-semibold tracking-tight text-slate-100">Template Viewer-Editor</h1>
          <p className="text-sm text-slate-300">
            {viewOnly
              ? 'Review workbook content in a read-first flow.'
              : 'Edit mapped fields safely while preserving workbook structure and formatting.'}
          </p>
        </div>
        <div className="rounded-xl border border-slate-700 bg-slate-800 px-3 py-1.5 text-sm text-slate-300">
          File: <span className="font-medium text-slate-100">{fileName || 'No file selected'}</span>
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-2">
        <button
          type="button"
          disabled={busy || !supportsFs}
          onClick={() => runAction(viewOnly ? () => openFileWithMode('VIEW') : openFile)}
          className="rounded-xl bg-sky-600 px-3 py-2 text-sm font-semibold text-white hover:bg-sky-500 disabled:cursor-not-allowed disabled:opacity-50"
        >
          Open
        </button>

        {!viewOnly && (
          <>
            <button
              type="button"
              disabled={busy || !canSameFileSave || !dirty}
              onClick={() => runAction(save)}
              className="rounded-xl bg-blue-600 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-50"
              title="Ctrl/Cmd+S"
            >
              Save
            </button>
            <button
              type="button"
              disabled={busy || !supportsFs || !hasWorkbook}
              onClick={() => runAction(saveAs)}
              className="rounded-xl bg-slate-800 px-3 py-2 text-sm font-semibold text-slate-100 ring-1 ring-slate-600 hover:bg-slate-700 disabled:cursor-not-allowed disabled:opacity-50"
            >
              Save as
            </button>
            <button
              type="button"
              disabled={busy || !hasWorkbook}
              onClick={() => runAction(downloadCopy)}
              className="rounded-xl bg-slate-800 px-3 py-2 text-sm font-semibold text-slate-100 ring-1 ring-slate-600 hover:bg-slate-700 disabled:cursor-not-allowed disabled:opacity-50"
            >
              Download copy
            </button>

            <span
              className={`ml-auto rounded-full px-3 py-1 text-xs font-semibold ${
                dirty ? 'bg-amber-500/20 text-amber-200' : 'bg-emerald-500/20 text-emerald-200'
              }`}
            >
              {dirty ? 'Unsaved changes' : 'Saved'}
            </span>
          </>
        )}
      </div>

      {!viewOnly && lastSavedAt && !dirty && (
        <p className="mt-2 text-xs text-slate-400">Last saved at {formatSavedTime(lastSavedAt)}</p>
      )}

      {!supportsFs && (
        <p className="mt-2 text-sm text-amber-300">
          File System Access API not available. Use Chrome or Edge for same-file save.
        </p>
      )}
      {error && <p className="mt-2 text-sm text-rose-300">{error}</p>}
    </header>
  )
}
