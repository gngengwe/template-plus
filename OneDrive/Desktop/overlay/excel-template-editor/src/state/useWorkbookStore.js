import { saveAs } from 'file-saver'
import { useSyncExternalStore } from 'react'
import {
  openXlsxFileHandle,
  readFileFromHandle,
  saveArrayBufferToHandle,
  saveAsArrayBuffer,
} from '../excel/fileSystem'
import { getSheet, readWorkbook, setCellByA1 } from '../excel/workbook'
import { patchWorkbookArrayBuffer } from '../excel/patchWorkbook'

const state = {
  workbook: null,
  activeSheetName: '',
  mode: 'AUTO',
  dirty: false,
  fileHandle: null,
  fileName: '',
  originalArrayBuffer: null,
  editsBySheet: {},
  lastSavedAt: null,
}

const listeners = new Set()
let snapshot = null

function emit() {
  listeners.forEach((listener) => listener())
}

function setState(patch) {
  Object.assign(state, patch)
  snapshot = { ...state, ...actions }
  emit()
}

async function buildPatchedBytes() {
  if (!state.workbook) throw new Error('Open an Excel workbook first.')
  if (!state.originalArrayBuffer) throw new Error('Original workbook bytes are unavailable.')

  if (!state.dirty) {
    return state.originalArrayBuffer
  }

  return patchWorkbookArrayBuffer(state.originalArrayBuffer, state.workbook, state.editsBySheet)
}

const actions = {
  async openFileWithMode(openMode = 'AUTO') {
    if (state.dirty && typeof window !== 'undefined') {
      const proceed = window.confirm('You have unsaved changes. Open another file and discard current edits?')
      if (!proceed) return
    }

    const handle = await openXlsxFileHandle()
    const buffer = await readFileFromHandle(handle)
    const workbook = readWorkbook(buffer)
    const fallbackSheet = workbook.SheetNames?.[0] ?? ''
    const defaultSheet = workbook.SheetNames.includes('Matter Details') ? 'Matter Details' : fallbackSheet

    setState({
      workbook,
      activeSheetName: defaultSheet,
      mode: openMode,
      dirty: false,
      fileHandle: handle,
      fileName: handle.name || 'workbook.xlsx',
      originalArrayBuffer: buffer,
      editsBySheet: {},
      lastSavedAt: null,
    })
  },

  async openFile() {
    await actions.openFileWithMode('AUTO')
  },

  setActiveSheet(name) {
    setState({ activeSheetName: name })
  },

  setMode(mode) {
    setState({ mode })
  },

  setCell(sheetName, a1, value) {
    if (!state.workbook) return
    const ws = getSheet(state.workbook, sheetName)
    if (!ws) return

    const updatedA1 = setCellByA1(ws, a1, value) || a1
    const cell = ws[updatedA1] || {}

    const prevSheetEdits = state.editsBySheet[sheetName] || {}
    const nextSheetEdits = {
      ...prevSheetEdits,
      [updatedA1]: {
        v: cell.v ?? '',
        t: cell.t ?? 's',
        styleKey: sheetName === 'Actions' ? 'actionsEntry' : null,
      },
    }

    setState({
      dirty: true,
      editsBySheet: {
        ...state.editsBySheet,
        [sheetName]: nextSheetEdits,
      },
    })
  },

  async save() {
    if (!state.fileHandle) {
      throw new Error('No file handle is available for same-file save. Use Open first.')
    }

    const bytes = await buildPatchedBytes()
    await saveArrayBufferToHandle(state.fileHandle, bytes)

    setState({
      dirty: false,
      editsBySheet: {},
      originalArrayBuffer: bytes,
      lastSavedAt: new Date().toISOString(),
    })
  },

  async saveAs() {
    const bytes = await buildPatchedBytes()
    const suggestedName = state.fileName || 'template-copy.xlsx'
    const handle = await saveAsArrayBuffer(bytes, suggestedName)

    setState({
      fileHandle: handle,
      fileName: handle.name || suggestedName,
      dirty: false,
      editsBySheet: {},
      originalArrayBuffer: bytes,
      lastSavedAt: new Date().toISOString(),
    })
  },

  async downloadCopy() {
    const bytes = await buildPatchedBytes()

    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })

    const nameBase = (state.fileName || 'template').replace(/\.xlsx$/i, '')
    saveAs(blob, `${nameBase}-copy.xlsx`)
  },
}

const store = {
  subscribe(listener) {
    listeners.add(listener)
    return () => listeners.delete(listener)
  },
  getSnapshot() {
    if (!snapshot) {
      snapshot = { ...state, ...actions }
    }
    return snapshot
  },
}

export function useWorkbookStore(selector = (s) => s) {
  const current = useSyncExternalStore(store.subscribe, store.getSnapshot)
  return selector(current)
}
