import {
  contentForInput,
  EN_US,
  MAX_COLS,
  MAX_ROWS,
  parseDateTime,
  validDateSerial,
  withRegional,
  Workbook,
  type CellData,
  type Regional,
  type StyleTable,
} from "@cells/core"
import type { OpenOptions } from "../xlsx/read"
import { DEFAULT_VIEW, type SheetView } from "../xlsx/sheet-read"
import { decodeText, type TextEncoding } from "./encoding"
import { CSV, detectDelimiter, parseDelimited, splitFixed, textLines } from "./split"

/** The Text Import Wizard's choices and column formats live in core, with the query tables that keep them. */
export type { ColumnFormat, TextImport } from "@cells/core"
import type { ColumnFormat, TextImport } from "@cells/core"

export const DATE_ORDERS = ["MDY", "DMY", "YMD", "MYD", "DYM", "YDM"] as const

export type TextRead = {
  workbook: Workbook
  views: SheetView[]
  activeSheet: number
  warnings: string[]
  /** How the file was written, so Save can write it back the same way. */
  encoding: TextEncoding
  bom: boolean
  delimiter: string
}

/** What opening a .csv does: comma-delimited (or the delimiter the file plainly uses), every column General. */
export function defaultImport(text: string, fileName: string): TextImport {
  const ext = /\.([^.]+)$/.exec(fileName)?.[1]?.toLowerCase()
  // A semicolon CSV (a European list separator) opens in columns too, where Excel's en-US default would put each
  // line into one cell.
  const delimiter = ext === "csv" ? detectDelimiter(text, ",") : detectDelimiter(text, "\t")
  return {
    kind: "delimited",
    delimiters: [delimiter],
    qualifier: '"',
    consecutive: false,
    breaks: [],
    startRow: 1,
    columns: [],
    decimal: ".",
    thousands: ",",
  }
}

/** The fields a text gives with these choices (all rows, or the first `limit` for a preview). */
export function splitText(text: string, o: TextImport, limit = Infinity): string[][] {
  const skip = Math.max(0, o.startRow - 1)
  if (o.kind === "fixed") {
    const lines = textLines(text, limit + skip).slice(skip)
    return splitFixed(lines, o.breaks)
  }
  return parseDelimited(text, o.delimiters.length ? o : CSV, limit + skip).slice(skip)
}

/** Opens a CSV or text file as a workbook of one sheet named after the file, as Excel does. */
export async function openText(
  bytes: Uint8Array,
  fileName: string,
  options: Partial<TextImport> & OpenOptions = {},
): Promise<TextRead> {
  const decoded = decodeText(bytes, options.encoding)
  const o: TextImport = { ...defaultImport(decoded.text, fileName), ...stripUndefined(options) }
  const rows = splitText(decoded.text, o)
  const warnings: string[] = []
  const wb = new Workbook([sheetNameFor(fileName)])
  const sheet = wb.sheets[0]
  const { tooWide, formulas } = await placeCells(
    wb.styles,
    rows,
    o,
    (r, c, data) => sheet.cells.set(r, c, data),
    options.progress,
  )
  if (rows.length > MAX_ROWS)
    warnings.push(
      `The file has ${rows.length.toLocaleString("en-US")} rows; only the first ${MAX_ROWS.toLocaleString("en-US")} fit on a sheet. Saving it over the file would lose the rest.`,
    )
  if (tooWide)
    warnings.push(
      `Some rows have more than ${MAX_COLS.toLocaleString("en-US")} fields; the ones past column XFD were left out.`,
    )
  wb.calc.rebuild()
  if (formulas && options.calculate !== false) wb.calc.calculateNow(true)
  return {
    workbook: wb,
    views: [{ ...DEFAULT_VIEW }],
    activeSheet: 0,
    warnings,
    encoding: decoded.encoding,
    bom: decoded.bom,
    delimiter: o.kind === "delimited" ? (o.delimiters[0] ?? ",") : "\t",
  }
}

/**
 * Puts the fields into cells (through `put`), each converted by its column's format; columns marked skip are left
 * out and the ones after move left. Stops at the sheet's last row and column.
 */
export async function placeCells(
  styles: StyleTable,
  rows: readonly (readonly string[])[],
  o: Pick<TextImport, "columns" | "decimal" | "thousands">,
  put: (row: number, col: number, data: CellData) => void,
  progress?: (done: number) => void | Promise<void>,
): Promise<{ tooWide: boolean; formulas: boolean }> {
  const textStyle = styles.intern({ numberFormat: "@" })
  const dateStyles = new Map<string, number>()
  const dateStyle = (code: string) => {
    let id = dateStyles.get(code)
    if (id === undefined) dateStyles.set(code, (id = styles.intern({ numberFormat: code })))
    return id
  }
  let tooWide = false
  let formulas = false
  const total = Math.min(rows.length, MAX_ROWS)
  for (let r = 0; r < total; r++) {
    const fields = rows[r]
    let col = 0
    for (let f = 0; f < fields.length; f++) {
      const format = o.columns[f] ?? "general"
      if (format === "skip") continue
      if (col >= MAX_COLS) {
        tooWide = true
        break
      }
      const data = cellFor(styles, fields[f], format, o, textStyle, dateStyle)
      if (data) {
        if (data.input?.startsWith("=")) formulas = true
        put(r, col, data)
      }
      col++
    }
    if (progress && r % 20000 === 19999) await progress(r / total)
  }
  return { tooWide, formulas }
}

/** One field as a cell: parsed as if typed (General), kept as text, or read as a date in the given order. */
export function cellFor(
  styles: StyleTable,
  field: string,
  format: ColumnFormat,
  o: Pick<TextImport, "decimal" | "thousands">,
  textStyle: number,
  dateStyle: (code: string) => number,
  /**
   * How General fields read dates: en-US for files (so a CSV opens the same everywhere, and the same in the worker),
   * the person's region for Text to Columns, which splits what the formula bar shows.
   */
  region: Regional = EN_US,
): CellData | undefined {
  if (field === "") return undefined
  if (format === "text") return { value: field, style: textStyle }
  if (format !== "general" && format !== "skip") {
    const date = dateIn(field, format)
    if (date) return { value: date.value, style: dateStyle(date.numberFormat) }
    return { value: field }
  }
  let text = field
  // Excel opens "@SUM(A1)" from a text file as a formula.
  if (/^@[A-Za-z_][\w.]*\(/.test(text)) text = "=" + text.slice(1)
  if (o.decimal !== "." || o.thousands !== ",") text = withUsSeparators(text, o.decimal, o.thousands)
  // The separators are en-US now; the date order and separator stay the region's.
  const reading: Regional = { ...region, decimal: ".", group: "," }
  return withRegional(reading, () => contentForInput(styles, undefined, text))
}

/** "1.234,5" with decimal "," and thousands "." as "1,234.5", so General reads it; other text is left alone. */
function withUsSeparators(text: string, decimal: string, thousands: string): string {
  const t = text.trim()
  const body = t.replace(/^[+\-(]/, "").replace(/[%)]$/, "")
  const allowed = new Set([...("0123456789" + decimal + thousands)])
  if (!body || ![...body].every((c) => allowed.has(c))) return text
  const parts = t.split(decimal)
  if (parts.length > 2) return text
  const whole = parts[0].split(thousands).join("")
  return parts.length === 2 ? `${whole}.${parts[1]}` : whole
}

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]

/** A date column's text ("31.12.2026", "2026-12-31 08:30") in the given order of day, month and year. */
export function dateIn(text: string, order: string): { value: number; numberFormat: string } | null {
  const m = /^\s*([0-9A-Za-z]+)[/\-. ]([0-9A-Za-z]+)[/\-. ]([0-9A-Za-z]+)(?:[ T]+(.+?))?\s*$/.exec(text)
  if (!m) return null
  const parts: Record<string, string> = {}
  for (let i = 0; i < 3; i++) parts[order[i]] = m[i + 1]
  const month = /^\d+$/.test(parts.M) ? Number(parts.M) : MONTHS.indexOf(parts.M.slice(0, 3).toLowerCase()) + 1
  if (!/^\d+$/.test(parts.D) || !/^\d+$/.test(parts.Y) || month < 1) return null
  let year = Number(parts.Y)
  // Two-digit years: 00–29 are 2000s, 30–99 the 1900s (Windows' and Excel's cutoff).
  if (parts.Y.length <= 2) year += year < 30 ? 2000 : 1900
  const serial = validDateSerial(year, month, Number(parts.D))
  if (serial === null) return null
  if (!m[4]) return { value: serial, numberFormat: "m/d/yyyy" }
  const time = parseDateTime(m[4])
  if (!time || time.value >= 1) return null
  return { value: serial + time.value, numberFormat: "m/d/yyyy h:mm" }
}

/** A sheet name from a file name: no extension, no characters a sheet name can't have, at most 31 characters. */
export function sheetNameFor(fileName: string): string {
  const name = fileName
    .replace(/\.[^.]+$/, "")
    .replace(/[[\]:*?/\\]/g, "_")
    .replace(/^'+|'+$/g, "")
    .slice(0, 31)
  return name.trim() || "Sheet1"
}

function stripUndefined<T extends object>(o: T): Partial<T> {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as Partial<T>
}
