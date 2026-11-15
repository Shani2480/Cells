import {
  cellName,
  contentForInput,
  currentRegion,
  editTextFor,
  regional,
  type CellAddress,
  type CellWrite,
  type Op,
  type Range,
  type Sheet,
  type StyleTable,
} from "@cells/core"
import { cellFor, splitText, type TextImport } from "@cells/io"

/** Text to Columns' first choices, as Excel's: delimited by Tab, quotes as the text qualifier. */
export const TEXT_COLUMNS_DEFAULT: TextImport = {
  kind: "delimited",
  delimiters: ["\t"],
  qualifier: '"',
  consecutive: false,
  breaks: [],
  startRow: 1,
  columns: [],
  decimal: ".",
  thousands: ",",
}

export const ONE_COLUMN_ONLY =
  "Cells can convert only one column at a time. The range can be many rows tall but no more than one column wide. Try again by selecting cells in one column only."
export const NOTHING_TO_PARSE = "No data was selected to parse."
export const REPLACE_DESTINATION = "Do you want to replace the contents of the destination cells?"

/** The text Text to Columns splits in each cell: what the formula bar shows (a formula's text, not its result). */
export function columnTexts(sheet: Sheet, styles: StyleTable, range: Range): string[] {
  const out: string[] = []
  for (let r = range.r0; r <= range.r1; r++) {
    const data = sheet.cells.get(r, range.c0)
    if (!data || data.value === undefined || data.value === null) out.push(data?.input ?? "")
    else out.push(data.input ?? editTextFor(data.value, styles.get(sheet.styleAt(r, range.c0)).numberFormat))
  }
  return out
}

/** One cell's fields. Fixed-width fields lose the spaces around them, as in Excel. */
export function splitCell(text: string, o: TextImport): string[] {
  if (text === "") return []
  const fields = splitText(text, { ...o, startRow: 1, columns: [] }, 1)[0] ?? []
  return o.kind === "fixed" ? fields.map((f) => f.trim()) : fields
}

/** The range Text to Columns works on: the selection, or the active cell's column of its region. */
export function textColumnsRange(sheet: Sheet, selected: Range, active: CellAddress): Range {
  if (selected.r0 !== selected.r1 || selected.c0 !== selected.c1) return selected
  const region = currentRegion(sheet, active)
  return { r0: region.r0, r1: region.r1, c0: active.col, c1: active.col }
}

/**
 * The writes of Text to Columns: each cell's fields go into the cells from `dest` rightwards, converted by their
 * column's format; skipped columns are left out. Every row clears as many columns as the widest row fills (Excel
 * does), and a cell that doesn't split, written back to itself, stays as it is.
 */
export function textToColumnsOps(
  sheetIndex: number,
  sheet: Sheet,
  styles: StyleTable,
  range: Range,
  o: TextImport,
  dest: CellAddress,
): { ops: Op[]; overwrites: boolean } {
  const texts = columnTexts(sheet, styles, range)
  const kept = (fields: string[]) => fields.filter((_, f) => (o.columns[f] ?? "general") !== "skip")
  const rows = texts.map((t) => splitCell(t, o))
  const width = Math.max(0, ...rows.map((r) => kept(r).length), ...rows.map((r) => (r.length ? 1 : 0)))
  const textStyle = styles.intern({ numberFormat: "@" })
  const dateStyle = (code: string) => styles.intern({ numberFormat: code })
  const cells: CellWrite[] = []
  let overwrites = false
  const inSource = (row: number, col: number) =>
    row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1
  rows.forEach((fields, i) => {
    const row = dest.row + i
    const sourceRow = range.r0 + i
    const formats = fields.map((_, f) => o.columns[f] ?? "general")
    const out = fields.filter((_, f) => formats[f] !== "skip")
    const outFormats = formats.filter((f) => f !== "skip")
    for (let j = 0; j < width; j++) {
      const col = dest.col + j
      const current = sheet.cells.get(row, col)
      if (row === sourceRow && col === range.c0 && fields.length === 1 && fields[0] === texts[i] && out.length === 1)
        continue
      if (current?.value !== undefined && current.value !== null && !inSource(row, col)) overwrites = true
      const field = out[j]
      let data
      if (field === undefined || field === "") data = current?.style ? { style: current.style } : undefined
      else if (outFormats[j] === "general" && field.startsWith("'")) data = { value: field, style: current?.style }
      else if (outFormats[j] === "general") {
        data = contentForInput(styles, current, field.trim())
        if (field.trim() === "") data = undefined
      } else data = cellFor(styles, field, outFormats[j], o, textStyle, dateStyle, regional())
      if (data && data.style === undefined) delete data.style
      cells.push({ row, col, data })
    }
  })
  return { ops: cells.length ? [{ kind: "setCells", sheet: sheetIndex, cells }] : [], overwrites }
}

/** "$B$2", the Destination box's text for a cell. */
export const absoluteName = (at: CellAddress) => cellName(at).replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2")
