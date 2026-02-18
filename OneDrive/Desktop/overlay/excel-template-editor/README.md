# Template Viewer-Editor

A React + Vite web app for editing a local Excel template through a clean overlay UI while preserving workbook formatting, styles, and merges.

## Stack

- React (JavaScript)
- Vite
- Tailwind CSS
- SheetJS (`xlsx`)
- Handsontable (`@handsontable/react`)
- `file-saver`

## Setup

```bash
npm install
npm run dev
```

Build for production:

```bash
npm run build
npm run preview
```

## Browser Support

- Same-file save requires File System Access API (`showOpenFilePicker`, `createWritable`) and is intended for Chrome/Edge.
- If File System Access API is unavailable, use **Download Copy**.

## Open/Save Model

- **Open** uses `window.showOpenFilePicker()` (not file input).
- **Save** writes back to the exact same opened `.xlsx` file handle via `handle.createWritable()`.
- **Save As** uses `showSaveFilePicker()` and updates the current handle.
- **Download Copy** exports a new workbook file using `file-saver`.
- Keyboard shortcut: **Ctrl/Cmd+S** triggers same-file Save when available.
- Unsaved-change guard: browser warns before page unload and confirms before opening a new file over unsaved edits.

## Excel Preservation Rules

The app keeps the original workbook object in memory and updates only target cell values:

- Read: `XLSX.read(arrayBuffer, { type: "array", cellDates: true, cellStyles: true })`
- Write: `XLSX.write(workbook, { bookType: "xlsx", type: "array" })`
- Edits: direct `ws[a1].v` and `ws[a1].t` updates only.

It does not regenerate sheets from JSON/AOA when saving, so existing formatting/merges/styles remain intact.

## Sheets and Modes

Expected workbook sheets:

1. `Matter Details`
2. `Inventor Information - COMPLETE`
3. `Actions`

Mode toggle options:

- `AUTO` (default):
  - Matter Details -> FORM
  - Inventor Information - COMPLETE -> FORM
  - Actions -> OVERLAY
  - Other sheets -> GRID
- `FORM`
- `GRID`

## Actions Overlay Stage Model

The `Actions` sheet is treated as a matrix:

- Column A = prompt labels
- Columns B..H = stages
  - B Engagement
  - C Disclosure: Pre-Search
  - D Disclosure: Post Search
  - E Restriction
  - F Office Action / Appeal Decision
  - G Advisory/AFCP/Pre-Appeal/Examiner's Answer
  - H Notice of Allowance

Sections:

- Event Metadata: rows 2-4
- Main Strategy: rows 6-17
- Continuation Proposal: rows 18-28 (shown by default for Notice of Allowance, optional toggle for others)

Extra features:

- Rating pills for rows 9-11 and 25-27 storing values as `High — <text>`, `Med — <text>`, or `Low — <text>`.
- Copy-from-stage action for rows 6-17.
- Jump-to-section sidebar.

## Mapping Adjustments

Mappings live in `src/excel/mapping.js`:

- Matter Details field-to-cell map
- Inventor field offsets
- Action stage columns and section row ranges

To adjust behavior, update mappings there and UI components will follow.

## Inventor Block Offsets

Inventor blocks repeat every 17 rows on `Inventor Information - COMPLETE`.

- Inventor 1 start row = 3 (`A3` label)
- For inventor `i` (0-based index), start row = `3 + i * 17`
- Value cells are in column `B` using configured offsets

The UI supports Inventor 1 through Inventor 8 and includes per-inventor clear actions.


