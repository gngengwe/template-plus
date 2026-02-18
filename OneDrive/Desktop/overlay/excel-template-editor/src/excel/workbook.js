import * as XLSX from 'xlsx'

export function readWorkbook(arrayBuffer) {
  return XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellStyles: true })
}

export function writeWorkbook(workbook) {
  return XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
    cellStyles: true,
    cellDates: true,
  })
}

export function getSheet(workbook, name) {
  return workbook?.Sheets?.[name]
}

function resolveMergedAnchorA1(ws, a1) {
  if (!ws?.['!merges']?.length) return a1
  const addr = XLSX.utils.decode_cell(a1)
  const containing = ws['!merges'].find(
    (m) =>
      addr.r >= m.s.r &&
      addr.r <= m.e.r &&
      addr.c >= m.s.c &&
      addr.c <= m.e.c,
  )
  if (!containing) return a1
  return XLSX.utils.encode_cell(containing.s)
}

export function getCellByA1(ws, a1) {
  const targetA1 = resolveMergedAnchorA1(ws, a1)
  return ws?.[targetA1]?.v ?? ''
}

function inferCellType(value) {
  if (value instanceof Date) return 'd'
  if (typeof value === 'number' && Number.isFinite(value)) return 'n'
  if (typeof value === 'boolean') return 'b'
  return 's'
}

function excelDateSerial(dateValue) {
  const date = dateValue instanceof Date ? dateValue : new Date(`${dateValue}T00:00:00`)
  if (Number.isNaN(date.getTime())) return null
  return 25569 + date.getTime() / 86400000
}

function hasDateLikeFormat(cell) {
  if (!cell?.z || typeof cell.z !== 'string') return false
  return /[ymd]/i.test(cell.z)
}

export function setCellByA1(ws, a1, value) {
  if (!ws) return null
  const targetA1 = resolveMergedAnchorA1(ws, a1)
  if (!ws[targetA1]) ws[targetA1] = {}

  const cell = ws[targetA1]

  delete cell.w
  delete cell.f

  if (value === '' || value === null || value === undefined) {
    cell.t = 'z'
    delete cell.v
    return targetA1
  }

  if ((value instanceof Date || typeof value === 'string') && hasDateLikeFormat(cell)) {
    const serial = excelDateSerial(value)
    if (serial !== null) {
      cell.v = serial
      cell.t = 'n'
      return targetA1
    }
  }

  cell.v = value
  cell.t = inferCellType(value)
  return targetA1
}
