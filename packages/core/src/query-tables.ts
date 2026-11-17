import type { Range } from "./address"
import { shiftRange } from "./conditional"

/**
 * Data from a text file that stays linked to it (Excel's text connection and its query table): From Text/CSV
 * records the file and the Text Import Wizard's choices, so Refresh reads the file again into the same range.
 */

/** Text encodings, with Excel's "File origin" code pages. */
export type TextEncoding = "utf-8" | "windows-1252" | "utf-16le" | "utf-16be"

/** A column's format in the Text Import Wizard: General, Text, a date order, or "Do not import column (skip)". */
export type ColumnFormat = "general" | "text" | "skip" | "MDY" | "DMY" | "YMD" | "MYD" | "DYM" | "YDM"

/** How a text file is read: the Text Import Wizard's choices (plain data, so it can go to the worker). */
export type TextImport = {
  /** Detected when absent. */
  encoding?: TextEncoding
  kind: "delimited" | "fixed"
  delimiters: string[]
  qualifier: string | null
  consecutive: boolean
  /** Fixed width: where each column starts (character positions, ascending). */
  breaks: number[]
  /** The first row imported, from 1. */
  startRow: number
  /** Per column; columns past the end are General. */
  columns: ColumnFormat[]
  /** Advanced: the decimal and thousands separators numbers use. */
  decimal: string
  thousands: string
}

/** One range of a sheet filled from a text file, and its connection. */
export type QueryTable = {
  /** The range's name (a sheet-level defined name in Excel files), unique on the sheet. */
  name: string
  /** The connection's name, shown in Queries & Connections; unique in the workbook. */
  connection: string
  description: string
  /** Where the data is: its last refresh filled this range. */
  range: Range
  /** The file: a path (desktop), a content URI (Android) or a file name (a browser). */
  source: string
  options: TextImport
  /** Refresh data when opening the file. */
  refreshOnLoad: boolean
}

export function copyTextImport(o: TextImport): TextImport {
  return { ...o, delimiters: [...o.delimiters], breaks: [...o.breaks], columns: [...o.columns] }
}

export function copyQueryTables(list: readonly QueryTable[]): QueryTable[] {
  return list.map((q) => ({ ...q, range: { ...q.range }, options: copyTextImport(q.options) }))
}

/** Rows or columns inserted or deleted: each range moves with its cells; one deleted whole goes. */
export function shiftQueryTables(list: QueryTable[], axis: "rows" | "cols", at: number, count: number): QueryTable[] {
  return list.flatMap((q) => {
    const range = shiftRange(q.range, axis, at, count)
    return range ? [{ ...q, range }] : []
  })
}

/** A name Excel accepts for a range or connection, from a file name ("Sales 2024.csv" → "Sales_2024"). */
export function queryTableName(fileName: string, taken: readonly string[]): string {
  let base = fileName
    .replace(/\.[^.]+$/, "")
    .replace(/[^A-Za-z0-9_.]/g, "_")
    .replace(/^[^A-Za-z_]/, (c) => `_${c}`)
  if (!base || /^[A-Za-z]{1,3}\d+$/.test(base) || /^[rc]$/i.test(base)) base = `_${base}`
  const used = new Set(taken.map((t) => t.toLowerCase()))
  if (!used.has(base.toLowerCase())) return base
  for (let n = 1; ; n++) if (!used.has(`${base}_${n}`.toLowerCase())) return `${base}_${n}`
}

/** Every query table of the workbook, with its sheet. */
export function allQueryTables(sheets: readonly { queryTables: QueryTable[] }[]) {
  return sheets.flatMap((sheet, index) => sheet.queryTables.map((table) => ({ sheet: index, table })))
}
