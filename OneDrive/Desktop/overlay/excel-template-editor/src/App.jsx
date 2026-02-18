import { useEffect, useMemo, useState } from 'react'
import Toolbar from './components/Toolbar'
import SheetTabs from './components/SheetTabs'
import ModeToggle from './components/ModeToggle'
import MatterDetailsForm from './components/MatterDetailsForm'
import InventorForm from './components/InventorForm'
import ActionsOverlay from './components/ActionsOverlay'
import GridSheet from './components/GridSheet'
import ViewSheet from './components/ViewSheet'
import ChangeSummary from './components/ChangeSummary'
import { useWorkbookStore } from './state/useWorkbookStore'

function resolveViewMode(mode, activeSheetName) {
  if (mode === 'VIEW') return 'VIEW'
  if (mode === 'FORM') return 'FORM'
  if (mode === 'GRID') return 'GRID'
  if (activeSheetName === 'Matter Details') return 'FORM'
  if (activeSheetName === 'Inventor Information - COMPLETE') return 'FORM'
  if (activeSheetName === 'Actions') return 'OVERLAY'
  return 'GRID'
}

function isViewOnlyExperience() {
  if (typeof window === 'undefined') return false
  const params = new URLSearchParams(window.location.search)
  return params.get('viewOnly') === '1' || window.location.pathname.toLowerCase().startsWith('/view')
}

export default function App() {
  const { workbook, fileHandle, activeSheetName, mode, dirty, save, setMode, openFileWithMode } = useWorkbookStore((state) => ({
    workbook: state.workbook,
    fileHandle: state.fileHandle,
    activeSheetName: state.activeSheetName,
    mode: state.mode,
    dirty: state.dirty,
    save: state.save,
    setMode: state.setMode,
    openFileWithMode: state.openFileWithMode,
  }))

  const [opening, setOpening] = useState(false)
  const [openError, setOpenError] = useState('')
  const viewOnly = useMemo(() => isViewOnlyExperience(), [])

  const sheetNames = workbook?.SheetNames ?? []
  const resolvedMode = resolveViewMode(mode, activeSheetName)
  const viewMode = viewOnly ? 'VIEW' : resolvedMode

  const content = useMemo(() => {
    if (!workbook || !activeSheetName) return null
    if (viewMode === 'VIEW') return <ViewSheet sheetName={activeSheetName} />
    if (viewMode === 'OVERLAY') return <ActionsOverlay sheetName={activeSheetName} />
    if (viewMode === 'FORM' && activeSheetName === 'Matter Details') return <MatterDetailsForm sheetName={activeSheetName} />
    if (viewMode === 'FORM' && activeSheetName === 'Inventor Information - COMPLETE') return <InventorForm sheetName={activeSheetName} />
    return <GridSheet sheetName={activeSheetName} />
  }, [workbook, activeSheetName, viewMode])

  useEffect(() => {
    const handleBeforeUnload = (event) => {
      if (!dirty) return
      event.preventDefault()
      event.returnValue = ''
    }
    window.addEventListener('beforeunload', handleBeforeUnload)
    return () => window.removeEventListener('beforeunload', handleBeforeUnload)
  }, [dirty])

  useEffect(() => {
    const onKeyDown = (event) => {
      const isSave = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's'
      if (!isSave) return
      if (viewOnly || !dirty || !workbook || !fileHandle) return
      event.preventDefault()
      save().catch(() => {})
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [viewOnly, dirty, workbook, fileHandle, save])

  const runOpen = async (openMode) => {
    setOpenError('')
    setOpening(true)
    try {
      await openFileWithMode(openMode)
    } catch (err) {
      setOpenError(err?.message || 'Unable to open workbook.')
    } finally {
      setOpening(false)
    }
  }

  return (
    <div className="min-h-screen bg-transparent px-4 py-8 md:px-10">
      <div className="mx-auto flex w-full max-w-7xl flex-col gap-5">
        <Toolbar viewOnly={viewOnly} />

        {!workbook ? (
          <section className="rounded-3xl border border-slate-700/80 bg-slate-900/85 p-8 shadow-glow backdrop-blur">
            <div className="mx-auto grid max-w-5xl grid-cols-1 gap-5 lg:grid-cols-2">
              <div className="rounded-2xl border border-slate-700 bg-slate-900/60 p-6">
                <h2 className="font-['Sora'] text-3xl font-semibold text-slate-100">Template Viewer-Editor</h2>
                <p className="mt-3 text-sm leading-6 text-slate-300">
                  Open a local Excel workbook and choose your workflow. View mode is optimized for long-scroll reading,
                  while edit mode uses mapped forms and structured overlays.
                </p>
                {viewOnly && (
                  <p className="mt-2 text-xs text-slate-400">
                    This is the view-only experience. Use <code className="rounded bg-slate-800 px-1">/</code> for full edit access.
                  </p>
                )}
                {openError && <p className="mt-3 text-sm text-rose-300">{openError}</p>}
              </div>

              <div className="rounded-2xl border border-slate-700 bg-slate-900/50 p-6">
                <p className="mb-4 text-sm font-semibold uppercase tracking-wide text-slate-400">Start</p>
                <div className="flex flex-col gap-3">
                  <button
                    type="button"
                    disabled={opening}
                    onClick={() => runOpen('VIEW')}
                    className="rounded-xl bg-sky-600 px-4 py-3 text-left text-white shadow-soft transition hover:bg-sky-500 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    <p className="text-sm font-semibold">Open and view</p>
                    <p className="text-xs text-sky-100">Read-focused, scroll-optimized layout.</p>
                  </button>

                  {!viewOnly && (
                    <button
                      type="button"
                      disabled={opening}
                      onClick={() => runOpen('AUTO')}
                      className="rounded-xl bg-slate-800 px-4 py-3 text-left text-slate-100 ring-1 ring-slate-600 transition hover:bg-slate-700 disabled:cursor-not-allowed disabled:opacity-60"
                    >
                      <p className="text-sm font-semibold">Open and edit</p>
                      <p className="text-xs text-slate-300">Mapped forms, actions overlay, and grid editing.</p>
                    </button>
                  )}
                </div>
              </div>
            </div>
          </section>
        ) : (
          <div className="rounded-3xl border border-slate-700/80 bg-slate-900/80 p-5 shadow-glow backdrop-blur">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
              <SheetTabs sheetNames={sheetNames} />
              {!viewOnly && <ModeToggle mode={mode} onModeChange={setMode} />}
            </div>
            {viewMode === 'VIEW' ? (
              <div>{content}</div>
            ) : (
              <div className="grid grid-cols-1 gap-4 xl:grid-cols-[1fr_320px]">
                <div>{content}</div>
                <ChangeSummary />
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
