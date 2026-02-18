export async function openXlsxFileHandle() {
  if (!('showOpenFilePicker' in window)) {
    throw new Error('File System Access API unavailable in this browser.')
  }

  const [handle] = await window.showOpenFilePicker({
    multiple: false,
    types: [
      {
        description: 'Excel Workbook',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
        },
      },
    ],
  })

  return handle
}

export async function readFileFromHandle(handle) {
  const file = await handle.getFile()
  return file.arrayBuffer()
}

export async function saveArrayBufferToHandle(handle, arrayBuffer) {
  if (!handle) throw new Error('No file handle available for save.')

  const permission = await handle.requestPermission({ mode: 'readwrite' })
  if (permission !== 'granted') throw new Error('Read/write permission denied for this file.')

  const writable = await handle.createWritable()
  await writable.write(arrayBuffer)
  await writable.close()
}

export async function saveAsArrayBuffer(arrayBuffer, suggestedName = 'template-copy.xlsx') {
  if (!('showSaveFilePicker' in window)) {
    throw new Error('showSaveFilePicker is unavailable in this browser.')
  }

  const handle = await window.showSaveFilePicker({
    suggestedName,
    types: [
      {
        description: 'Excel Workbook',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
        },
      },
    ],
  })

  await saveArrayBufferToHandle(handle, arrayBuffer)
  return handle
}

