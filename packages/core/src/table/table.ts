import type { Range } from "../address"
import { shiftRange } from "../conditional"
import { copyAutoFilter, type AutoFilter } from "../filter"
import type { Area } from "../formula/types"
import type { Sheet } from "../sheet"
import { numberToEditText, type ErrorCode } from "../values"

/**
 * Tables (Insert › Table, Home › Format as Table), as Excel's ListObject: a range with a header row, data rows and
 * an optional total row, named columns, a table style, and structured references (Table1[Column], [@Column]).
 */

/** The total row's function for a column (Excel's totalsRowFunction). */
export type TotalsFunction =
  "none" | "sum" | "min" | "max" | "average" | "count" | "countNums" | "stdDev" | "var" | "custom"

export type TableColumn = {
  name: string
  /** The total row: a function, or a label (text in the total row cell), or a formula of its own ("custom"). */
  totals?: TotalsFunction
  totalsLabel?: string
  /** A calculated column's formula (Cells' syntax, without "="), filled down the column. */
  formula?: string
}

export type Table = {
  /** Unique in the workbook (the file's table id). */
  id: number
  /** Unique among tables and defined names, as Excel requires. */
  name: string
  range: Range
  /** Header Row (on by default). */
  header: boolean
  /** Total Row. */
  totals: boolean
  columns: TableColumn[]
  /** A built-in style ("TableStyleMedium2") or none (null). */
  style: string | null
  firstColumn: boolean
  lastColumn: boolean
  rowStripes: boolean
  columnStripes: boolean
  /** Filter Button: the header row's dropdowns, with their filters and sort. */
  autoFilter: AutoFilter | null
  /** The query (Data › Get Data) that fills the table on Refresh. */
  query?: string
}

/** Excel's style for a new table. */
export const DEFAULT_TABLE_STYLE = "TableStyleMedium2"

export function copyTable(t: Table): Table {
  return {
    ...t,
    range: { ...t.range },
    columns: t.columns.map((c) => ({ ...c })),
    autoFilter: copyAutoFilter(t.autoFilter),
  }
}

export const copyTables = (tables: readonly Table[]): Table[] => tables.map(copyTable)

/** The table a cell of a sheet belongs to, or null. */
export function tableAt(sheet: Sheet, row: number, col: number): Table | null {
  for (const t of sheet.tables) {
    const r = t.range
    if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1) return t
  }
  return null
}

/** Every table of the workbook, with its sheet's index. */
export function allTables(sheets: readonly Sheet[]): { sheet: number; table: Table }[] {
  const out: { sheet: number; table: Table }[] = []
  sheets.forEach((s, i) => s.tables?.forEach((table) => out.push({ sheet: i, table })))
  return out
}

export function findTable(sheets: readonly Sheet[], name: string): { sheet: number; table: Table } | null {
  const upper = name.toUpperCase()
  return allTables(sheets).find((t) => t.table.name.toUpperCase() === upper) ?? null
}

/** The rows of each part of a table (null when the part is off). */
export function tableRows(t: Table): { header: number | null; data: [number, number]; totals: number | null } {
  const r = t.range
  return {
    header: t.header ? r.r0 : null,
    data: [r.r0 + (t.header ? 1 : 0), r.r1 - (t.totals ? 1 : 0)],
    totals: t.totals ? r.r1 : null,
  }
}

/** "Table1", "Table2"…: the first name no table or defined name has. */
export function nextTableName(sheets: readonly Sheet[], taken: readonly string[] = []): string {
  const used = new Set([
    ...allTables(sheets).map((t) => t.table.name.toUpperCase()),
    ...taken.map((n) => n.toUpperCase()),
  ])
  for (let i = 1; ; i++) if (!used.has(`TABLE${i}`)) return `Table${i}`
}

export function nextTableId(sheets: readonly Sheet[]): number {
  return allTables(sheets).reduce((m, t) => Math.max(m, t.table.id), 0) + 1
}

/** Column names made unique as Excel does (a repeat gets 2, 3…; a blank becomes Column1, Column2…). */
export function uniqueColumnNames(names: readonly string[]): string[] {
  const out: string[] = []
  const seen = new Set<string>()
  names.forEach((raw, i) => {
    let name = raw.replace(/[\r\n]+/g, " ").trim() ? raw : `Column${i + 1}`
    if (seen.has(name.toUpperCase())) {
      let n = 2
      while (seen.has(`${name}${n}`.toUpperCase())) n++
      name = `${name}${n}`
    }
    seen.add(name.toUpperCase())
    out.push(name)
  })
  return out
}

/** A header cell's text as a column name. */
function headerText(value: unknown): string {
  if (value === undefined || value === null) return ""
  if (typeof value === "number") return numberToEditText(value)
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE"
  if (typeof value === "object" && value && "error" in value) return String((value as { error: string }).error)
  return String(value)
}

/**
 * Column names follow the header cells written in `area` (made unique as Excel does). Returns whether a name
 * changed.
 */
export function syncTableHeaders(sheet: Sheet, area: Range): boolean {
  let changed = false
  sheet.tables = sheet.tables.map((t) => {
    const r = t.range
    if (!t.header || r.r0 < area.r0 || r.r0 > area.r1 || r.c1 < area.c0 || r.c0 > area.c1) return t
    const texts = t.columns.map((_, i) => headerText(sheet.cells.get(r.r0, r.c0 + i)?.value))
    const names = uniqueColumnNames(texts)
    if (names.every((n, i) => n === t.columns[i].name)) return t
    changed = true
    const next = copyTable(t)
    next.columns.forEach((c, i) => (c.name = names[i]))
    return next
  })
  return changed
}

// ── Rows and columns inserted or deleted ──────────────────────────────────────────────────────────────────────

/**
 * Tables follow inserted and deleted rows and columns. Columns inserted inside a table become new table columns
 * ("Column3"); deleted ones leave it. A table whose cells all go is dropped.
 */
export function shiftTables(sheet: Sheet, axis: "rows" | "cols", at: number, count: number) {
  if (!sheet.tables.length) return
  const next: Table[] = []
  for (const t of sheet.tables) {
    const range = shiftRange(t.range, axis, at, count)
    if (!range) continue
    const table = copyTable(t)
    table.range = range
    if (axis === "cols") {
      const c0 = t.range.c0
      if (count > 0 && at > c0 && at <= t.range.c1) {
        const names = table.columns.map((c) => c.name)
        const added: TableColumn[] = []
        for (let i = 0; i < count; i++) {
          let n = names.length + 1
          while (names.some((x) => x.toUpperCase() === `COLUMN${n}`)) n++
          names.push(`Column${n}`)
          added.push({ name: `Column${n}` })
        }
        table.columns.splice(at - c0, 0, ...added)
      } else if (count < 0) {
        const from = Math.max(at, c0) - c0
        const to = Math.min(at - count - 1, t.range.c1) - c0
        if (to >= from) table.columns.splice(from, to - from + 1)
      }
      if (table.columns.length !== range.c1 - range.c0 + 1) continue
    } else {
      // Rows deleted over the header or total row turn those parts off; a table needs a data row.
      if (count < 0) {
        const end = at - count - 1
        if (t.header && t.range.r0 >= at && t.range.r0 <= end) table.header = false
        if (t.totals && t.range.r1 >= at && t.range.r1 <= end) table.totals = false
        const rows = range.r1 - range.r0 + 1 - (table.header ? 1 : 0) - (table.totals ? 1 : 0)
        if (rows < 1) continue
      }
    }
    if (table.autoFilter) {
      const header = table.header ? table.range.r0 : null
      if (header === null) table.autoFilter = null
      else
        table.autoFilter = {
          ...table.autoFilter,
          range: { ...table.range, r1: table.range.r1 - (table.totals ? 1 : 0) },
        }
    }
    next.push(table)
  }
  sheet.tables = next
}

// ── Structured references ─────────────────────────────────────────────────────────────────────────────────────

/** A structured reference, read: which table, which parts (rows) and which columns. */
export type StructuredSpec = {
  /** null: the table the formula is in ([@Qty], [Qty]). */
  table: string | null
  headers: boolean
  data: boolean
  totals: boolean
  thisRow: boolean
  /** First and last column by name (null: every column). */
  c0: string | null
  c1: string | null
}

const KEYWORDS: Record<string, "all" | "data" | "headers" | "totals" | "thisRow"> = {
  "#ALL": "all",
  "#DATA": "data",
  "#HEADERS": "headers",
  "#TOTALS": "totals",
  "#THIS ROW": "thisRow",
}

/** The text between a "[" at `from` and its "]" (escapes kept), and where it ends. */
function group(s: string, from: number): { text: string; end: number } | null {
  let depth = 0
  for (let j = from; j < s.length; j++) {
    if (s[j] === "'") {
      j++
      continue
    }
    if (s[j] === "[") depth++
    else if (s[j] === "]" && --depth === 0) return { text: s.slice(from + 1, j), end: j + 1 }
  }
  return null
}

/** A column name with Excel's escapes ('[ '] '# '') taken out. */
const unescape = (s: string) => s.replace(/'(.)/g, "$1")

/** Reads `Table1[[#Headers],[Qty]:[Price]]`, `[@Qty]`, `Table1[]`… (null: not a structured reference). */
export function parseStructured(text: string): StructuredSpec | null {
  const open = text.indexOf("[")
  if (open < 0) return null
  const outer = group(text, open)
  if (!outer || outer.end !== text.length) return null
  const spec: StructuredSpec = {
    table: open > 0 ? text.slice(0, open) : null,
    headers: false,
    data: false,
    totals: false,
    thisRow: false,
    c0: null,
    c1: null,
  }
  const columns: string[] = []
  const keyword = (k: string) => {
    const kind = KEYWORDS[k.trim().toUpperCase().replace(/\s+/g, " ")]
    if (!kind) return false
    if (kind === "all") spec.headers = spec.data = spec.totals = true
    else spec[kind] = true
    return true
  }
  // A list of [..] groups, with ":" between two column groups.
  const groups = (s: string): boolean => {
    let i = 0
    let range = false
    while (i < s.length) {
      const ch = s[i]
      if (ch === " " || ch === ",") {
        i++
        continue
      }
      if (ch === ":") {
        range = true
        i++
        continue
      }
      if (ch !== "[") return false
      const g = group(s, i)
      if (!g) return false
      i = g.end
      if (g.text.startsWith("#")) {
        if (!keyword(g.text)) return false
      } else {
        if (range && columns.length !== 1) return false
        columns.push(unescape(g.text))
      }
    }
    return true
  }
  let inner = outer.text
  if (inner.startsWith("@")) {
    spec.thisRow = true
    inner = inner.slice(1).trim()
    if (inner && !inner.startsWith("[")) columns.push(unescape(inner))
    else if (inner && !groups(inner)) return null
  } else if (inner.trim().startsWith("[")) {
    if (!groups(inner.trim())) return null
  } else if (inner.trim().startsWith("#")) {
    if (!keyword(inner)) return null
  } else if (inner !== "") columns.push(unescape(inner))
  if (columns.length > 2) return null
  if (!spec.headers && !spec.totals && !spec.thisRow) spec.data = true
  if (columns.length) {
    spec.c0 = columns[0]
    spec.c1 = columns[columns.length - 1]
  }
  return spec
}

/** Whether a column name must be in its own brackets ([@[Unit Price]], Table1[[#Totals],[Q1]]). */
const needsBrackets = (name: string) => /[\s\t\n\r,:.[\]#'"{}$^&*+=\-<>/@!()%~`;?|\\]/.test(name)
const escape = (name: string) => name.replace(/['[\]#]/g, "'$&")

/**
 * Writes a structured reference. `own`: the formula is in this table, so the table's name is left out ([Qty]).
 * `file`: as files store it, always with the table's name and #This Row spelled out.
 */
export function printStructured(spec: StructuredSpec, tableName: string | null, own: boolean, file = false): string {
  const name = file || !own || !tableName ? (tableName ?? "") : ""
  const col = (c: string) => `[${escape(c)}]`
  const cols = spec.c0 === null ? [] : spec.c1 && spec.c1 !== spec.c0 ? [spec.c0, spec.c1] : [spec.c0]
  const colText = cols.map(col).join(":")
  if (spec.thisRow && !file) {
    if (!cols.length) return `${name}[@]`
    if (cols.length === 1 && !needsBrackets(cols[0])) return `${name}[@${escape(cols[0])}]`
    return `${name}[@${colText}]`
  }
  const parts: string[] = []
  if (spec.thisRow) parts.push("[#This Row]")
  else if (spec.headers && spec.data && spec.totals) parts.push("[#All]")
  else {
    if (spec.headers) parts.push("[#Headers]")
    if (spec.data && (spec.headers || spec.totals || !cols.length)) parts.push("[#Data]")
    if (spec.totals) parts.push("[#Totals]")
  }
  // Table1[] is the data; Table1[[#Data]] is written as Table1[#Data] (one item needs no inner brackets).
  if (!cols.length) {
    if (parts.length === 1) return `${name}${parts[0]}`
    return `${name}[${parts.join(",")}]`
  }
  if (!parts.length) {
    if (cols.length === 1 && !needsBrackets(cols[0])) return `${name}[${escape(cols[0])}]`
    return `${name}[${colText}]`
  }
  return `${name}[${[...parts, colText].join(",")}]`
}

/** The area a structured reference stands for, from a cell (`sheet`, `row`, `col`), or the error it gives. */
export function structuredArea(
  spec: StructuredSpec,
  tables: readonly { sheet: number; table: Table }[],
  sheet: number,
  row: number,
  col: number,
): Area | ErrorCode {
  const upper = spec.table?.toUpperCase()
  const found = upper
    ? tables.find((t) => t.table.name.toUpperCase() === upper)
    : tables.find(
        (t) =>
          t.sheet === sheet &&
          row >= t.table.range.r0 &&
          row <= t.table.range.r1 &&
          col >= t.table.range.c0 &&
          col <= t.table.range.c1,
      )
  if (!found) return "#REF!"
  const t = found.table
  const index = (name: string) => t.columns.findIndex((c) => c.name.toUpperCase() === name.toUpperCase())
  let c0 = t.range.c0
  let c1 = t.range.c1
  if (spec.c0 !== null) {
    const a = index(spec.c0)
    const b = index(spec.c1 ?? spec.c0)
    if (a < 0 || b < 0) return "#REF!"
    c0 = t.range.c0 + Math.min(a, b)
    c1 = t.range.c0 + Math.max(a, b)
  }
  const rows = tableRows(t)
  if (spec.thisRow) {
    // The formula's own row; outside the table's data rows there is none.
    if (found.sheet !== sheet || row < rows.data[0] || row > rows.data[1]) return "#VALUE!"
    return { sheet: found.sheet, r0: row, r1: row, c0, c1 }
  }
  const parts: [number, number][] = []
  if (spec.headers && rows.header !== null) parts.push([rows.header, rows.header])
  if (spec.data && rows.data[1] >= rows.data[0]) parts.push(rows.data)
  if (spec.totals && rows.totals !== null) parts.push([rows.totals, rows.totals])
  if (!parts.length) return "#REF!"
  return { sheet: found.sheet, r0: parts[0][0], r1: parts[parts.length - 1][1], c0, c1 }
}

/** Whether a range would make a valid table: inside the sheet, not overlapping another table. */
export function tableOverlaps(sheet: Sheet, range: Range, except?: Table): boolean {
  return sheet.tables.some(
    (t) =>
      t !== except &&
      t.range.r0 <= range.r1 &&
      t.range.r1 >= range.r0 &&
      t.range.c0 <= range.c1 &&
      t.range.c1 >= range.c0,
  )
}
