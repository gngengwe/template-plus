import JSZip from 'jszip'
import * as XLSX from 'xlsx'

const XML_NS = 'http://www.w3.org/XML/1998/namespace'

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, 'application/xml')
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc)
}

function normalizeSheetPath(relTarget) {
  if (!relTarget) return null
  const raw = relTarget.startsWith('/') ? relTarget.slice(1) : relTarget
  const base = raw.startsWith('xl/') ? raw : `xl/${raw.replace(/^\.\//, '')}`

  const parts = []
  base.split('/').forEach((part) => {
    if (!part || part === '.') return
    if (part === '..') {
      parts.pop()
      return
    }
    parts.push(part)
  })

  return parts.join('/')
}

function isBlankValue(value) {
  return value === '' || value === null || value === undefined
}

function getSheetPathMap(workbookXml, relsXml) {
  const workbookDoc = parseXml(workbookXml)
  const relsDoc = parseXml(relsXml)

  const relMap = new Map()
  Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((relNode) => {
    const id = relNode.getAttribute('Id')
    const target = relNode.getAttribute('Target')
    if (id && target) relMap.set(id, target)
  })

  const map = new Map()
  Array.from(workbookDoc.getElementsByTagName('sheet')).forEach((sheetNode) => {
    const name = sheetNode.getAttribute('name')
    const relId =
      sheetNode.getAttribute('r:id') ||
      sheetNode.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')

    if (!name || !relId) return
    const path = normalizeSheetPath(relMap.get(relId))
    if (path) map.set(name, path)
  })

  return map
}

function upsertRow(sheetDataNode, rowNumber) {
  const rows = Array.from(sheetDataNode.getElementsByTagName('row'))
  const existing = rows.find((row) => Number(row.getAttribute('r')) === rowNumber)
  if (existing) return existing

  const ns = sheetDataNode.namespaceURI || sheetDataNode.ownerDocument.documentElement?.namespaceURI || null
  const rowNode = sheetDataNode.ownerDocument.createElementNS(ns, 'row')
  rowNode.setAttribute('r', String(rowNumber))

  const insertBefore = rows.find((row) => Number(row.getAttribute('r')) > rowNumber)
  if (insertBefore) {
    sheetDataNode.insertBefore(rowNode, insertBefore)
  } else {
    sheetDataNode.appendChild(rowNode)
  }

  return rowNode
}

function upsertCell(rowNode, a1, colIndexZeroBased) {
  const cells = Array.from(rowNode.getElementsByTagName('c'))
  const existing = cells.find((cell) => cell.getAttribute('r') === a1)
  if (existing) return existing

  const ns = rowNode.namespaceURI || rowNode.ownerDocument.documentElement?.namespaceURI || null
  const node = rowNode.ownerDocument.createElementNS(ns, 'c')
  node.setAttribute('r', a1)

  const insertBefore = cells.find((cell) => {
    const ref = cell.getAttribute('r')
    if (!ref) return false
    const decoded = XLSX.utils.decode_cell(ref)
    return decoded.c > colIndexZeroBased
  })

  if (insertBefore) {
    rowNode.insertBefore(node, insertBefore)
  } else {
    rowNode.appendChild(node)
  }

  return node
}

function clearCellValueNodes(cellNode) {
  Array.from(cellNode.childNodes).forEach((node) => {
    if (node.nodeType === 1) {
      cellNode.removeChild(node)
    }
  })
}

function setBlank(cellNode) {
  clearCellValueNodes(cellNode)
  cellNode.removeAttribute('t')
}

function setBoolean(cellNode, value) {
  clearCellValueNodes(cellNode)
  cellNode.setAttribute('t', 'b')
  const ns = cellNode.namespaceURI || cellNode.ownerDocument.documentElement?.namespaceURI || null
  const vNode = cellNode.ownerDocument.createElementNS(ns, 'v')
  vNode.textContent = value ? '1' : '0'
  cellNode.appendChild(vNode)
}

function setNumber(cellNode, value) {
  clearCellValueNodes(cellNode)
  cellNode.removeAttribute('t')
  const ns = cellNode.namespaceURI || cellNode.ownerDocument.documentElement?.namespaceURI || null
  const vNode = cellNode.ownerDocument.createElementNS(ns, 'v')
  vNode.textContent = String(value)
  cellNode.appendChild(vNode)
}

function setInlineString(cellNode, value) {
  clearCellValueNodes(cellNode)
  cellNode.setAttribute('t', 'inlineStr')

  const ns = cellNode.namespaceURI || cellNode.ownerDocument.documentElement?.namespaceURI || null
  const isNode = cellNode.ownerDocument.createElementNS(ns, 'is')
  const tNode = cellNode.ownerDocument.createElementNS(ns, 't')
  const text = String(value)
  if (/^\s|\s$|\n/.test(text)) {
    tNode.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  }
  tNode.textContent = text
  isNode.appendChild(tNode)
  cellNode.appendChild(isNode)
}

function toExcelDateSerial(rawValue) {
  const d = rawValue instanceof Date ? rawValue : new Date(rawValue)
  if (Number.isNaN(d.getTime())) return null
  return 25569 + d.getTime() / 86400000
}

async function getStyleIds(zip) {
  const stylesEntry = zip.file('xl/styles.xml')
  if (!stylesEntry) return {}

  const stylesXml = await stylesEntry.async('string')
  const doc = parseXml(stylesXml)
  const styleSheet = doc.getElementsByTagName('styleSheet')[0]
  if (!styleSheet) return {}

  const ns = styleSheet.namespaceURI || doc.documentElement?.namespaceURI || null
  let cellXfs = doc.getElementsByTagName('cellXfs')[0]
  if (!cellXfs) {
    cellXfs = doc.createElementNS(ns, 'cellXfs')
    cellXfs.setAttribute('count', '0')
    styleSheet.appendChild(cellXfs)
  }

  const xfs = Array.from(cellXfs.getElementsByTagName('xf'))
  const existingIdx = xfs.findIndex((xf) => {
    const align = xf.getElementsByTagName('alignment')[0]
    if (!align) return false
    return (
      (align.getAttribute('horizontal') || '').toLowerCase() === 'general' &&
      (align.getAttribute('vertical') || '').toLowerCase() === 'top' &&
      align.getAttribute('wrapText') === '1'
    )
  })

  if (existingIdx >= 0) {
    return { actionsEntry: existingIdx }
  }

  const baseXf = xfs[0]
  const newXf = baseXf ? baseXf.cloneNode(true) : doc.createElementNS(ns, 'xf')
  if (!newXf.getAttribute('numFmtId')) newXf.setAttribute('numFmtId', '0')
  if (!newXf.getAttribute('fontId')) newXf.setAttribute('fontId', '0')
  if (!newXf.getAttribute('fillId')) newXf.setAttribute('fillId', '0')
  if (!newXf.getAttribute('borderId')) newXf.setAttribute('borderId', '0')
  if (!newXf.getAttribute('xfId')) newXf.setAttribute('xfId', '0')
  newXf.setAttribute('applyAlignment', '1')
  newXf.setAttribute('fontId', '0')

  Array.from(newXf.getElementsByTagName('alignment')).forEach((node) => newXf.removeChild(node))

  const alignment = doc.createElementNS(ns, 'alignment')
  alignment.setAttribute('horizontal', 'general')
  alignment.setAttribute('vertical', 'top')
  alignment.setAttribute('wrapText', '1')
  newXf.appendChild(alignment)

  cellXfs.appendChild(newXf)
  cellXfs.setAttribute('count', String(cellXfs.getElementsByTagName('xf').length))

  const newIdx = cellXfs.getElementsByTagName('xf').length - 1
  zip.file('xl/styles.xml', serializeXml(doc))
  return { actionsEntry: newIdx }
}

function patchCell(cellNode, edit, styleIds) {
  const value = edit?.v
  const type = edit?.t
  const styleKey = edit?.styleKey
  const styleId = styleKey ? styleIds?.[styleKey] : null

  if (styleId !== null && styleId !== undefined) {
    cellNode.setAttribute('s', String(styleId))
  }

  if (isBlankValue(value)) {
    setBlank(cellNode)
    return
  }

  if (type === 'b' || typeof value === 'boolean') {
    setBoolean(cellNode, Boolean(value))
    return
  }

  if (type === 'd') {
    const serial = toExcelDateSerial(value)
    if (serial !== null) {
      setNumber(cellNode, serial)
      return
    }
  }

  if (type === 'n' || typeof value === 'number') {
    const n = typeof value === 'number' ? value : Number(value)
    if (Number.isFinite(n)) {
      setNumber(cellNode, n)
      return
    }
  }

  setInlineString(cellNode, value)
}

function patchWorksheetXml(xmlText, editsByA1, styleIds) {
  const doc = parseXml(xmlText)
  const worksheet = doc.getElementsByTagName('worksheet')[0]
  if (!worksheet) return xmlText

  let sheetData = doc.getElementsByTagName('sheetData')[0]
  if (!sheetData) {
    const ns = worksheet.namespaceURI || doc.documentElement?.namespaceURI || null
    sheetData = doc.createElementNS(ns, 'sheetData')
    worksheet.appendChild(sheetData)
  }

  Object.entries(editsByA1).forEach(([a1, edit]) => {
    const decoded = XLSX.utils.decode_cell(a1)
    const rowNumber = decoded.r + 1
    const rowNode = upsertRow(sheetData, rowNumber)
    const cellNode = upsertCell(rowNode, a1, decoded.c)
    patchCell(cellNode, edit, styleIds)
  })

  return serializeXml(doc)
}

export async function patchWorkbookArrayBuffer(originalArrayBuffer, workbook, editsBySheet) {
  if (!originalArrayBuffer) throw new Error('Original workbook bytes are unavailable.')

  const zip = await JSZip.loadAsync(originalArrayBuffer)
  const workbookEntry = zip.file('xl/workbook.xml')
  const relsEntry = zip.file('xl/_rels/workbook.xml.rels')

  if (!workbookEntry || !relsEntry) {
    throw new Error('Workbook XML structure is missing expected files.')
  }

  const [workbookXml, relsXml] = await Promise.all([
    workbookEntry.async('string'),
    relsEntry.async('string'),
  ])

  const sheetPathMap = getSheetPathMap(workbookXml, relsXml)
  const styleIds = await getStyleIds(zip)

  let dirtyCount = 0
  let patchedCount = 0

  for (const [sheetName, editsByA1] of Object.entries(editsBySheet || {})) {
    const entries = Object.entries(editsByA1 || {})
    if (!entries.length) continue
    dirtyCount += entries.length

    let sheetPath = sheetPathMap.get(sheetName)
    if (!sheetPath) {
      const idx = workbook?.SheetNames?.indexOf(sheetName)
      if (idx >= 0) {
        const fallback = `xl/worksheets/sheet${idx + 1}.xml`
        if (zip.file(fallback)) sheetPath = fallback
      }
    }

    if (!sheetPath) {
      throw new Error(`Unable to resolve worksheet XML path for sheet "${sheetName}".`)
    }

    const sheetEntry = zip.file(sheetPath)
    if (!sheetEntry) {
      throw new Error(`Worksheet XML not found at path "${sheetPath}" for sheet "${sheetName}".`)
    }

    const xmlText = await sheetEntry.async('string')
    const patchedXml = patchWorksheetXml(xmlText, editsByA1, styleIds)
    zip.file(sheetPath, patchedXml)
    patchedCount += entries.length
  }

  if (dirtyCount > 0 && patchedCount === 0) {
    throw new Error('No edits were patched into workbook XML.')
  }

  return zip.generateAsync({ type: 'arraybuffer' })
}
