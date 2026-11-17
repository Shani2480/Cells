import { MAX_COLS } from "./address"
import type { CellValue } from "./values"

/**
 * Links to other workbooks: formulas such as ='[Price List.xlsx]Prices'!B2 or ='Price List.xlsx'!Rate. Like Excel
 * with the other workbook closed, Cells calculates them from the values last read from that file (kept in the
 * workbook, and in .xlsx files as an external link part), until Data › Workbook Links updates them from the file.
 */

/** One sheet of a linked workbook: the values formulas here read from it. */
export type LinkedSheet = {
  name: string
  /** row × 16384 + col → value. */
  cells: Map<number, CellValue>
}

export type ExternalBook = {
  /** The file's name as formulas write it ("Price List.xlsx"). */
  name: string
  /** Where the file is: a path relative to this workbook's folder (as Excel stores it), or absolute. */
  target: string
  /** The absolute path too, when known (Excel 2021's alternate URL). */
  absolute?: string
  sheets: LinkedSheet[]
  /** The linked workbook's names that formulas here use, with what they refer to there. */
  names: { name: string; refersTo: string; sheet?: number }[]
}

export const linkKey = (row: number, col: number) => row * MAX_COLS + col

export function copyExternalBooks(list: readonly ExternalBook[]): ExternalBook[] {
  return list.map((b) => ({
    ...b,
    sheets: b.sheets.map((s) => ({ name: s.name, cells: new Map(s.cells) })),
    names: b.names.map((n) => ({ ...n })),
  }))
}

/**
 * Pseudo sheet indexes for the engine: linked sheets sit far after the workbook's own (16384 + book × 256 + sheet),
 * and 255 is a book's own scope (for its names). They stay below what the engine's cell keys can hold.
 */
export const EXTERNAL_SHEET_BASE = 16384
export const BOOK_SCOPE = 255

export const externalSheetIndex = (book: number, sheet: number) => EXTERNAL_SHEET_BASE + book * 256 + sheet

export function externalSheetOf(index: number): { book: number; sheet: number } | null {
  if (index < EXTERNAL_SHEET_BASE) return null
  const n = index - EXTERNAL_SHEET_BASE
  return { book: Math.floor(n / 256), sheet: n % 256 }
}

/** "[Price List.xlsx]Prices" → its parts; a sheet name without a book gives null. */
export function splitBookSheet(name: string): { book: string; sheet: string | null } | null {
  const m = /^(?:.*[\\/])?\[([^\]]+)\](.*)$/.exec(name)
  if (!m) return null
  return { book: m[1], sheet: m[2] || null }
}

/** The file name a link is known by, from its path ("C:\Data\Price List.xlsx" → "Price List.xlsx"). */
export function bookNameOf(path: string): string {
  let decoded = path
  try {
    decoded = decodeURIComponent(path)
  } catch {
    // Not percent-encoded.
  }
  return (
    decoded
      .replace(/^file:\/+/i, "")
      .split(/[\\/]/)
      .pop() || decoded
  )
}

/** The pseudo sheet index of "[Book]Sheet" (or of "Book" alone, for a name in it); -1 when there's no such link. */
export function externalIndexOf(books: readonly ExternalBook[], name: string, sheetNames: readonly string[]): number {
  const parts = splitBookSheet(name)
  const upper = (s: string) => s.toUpperCase()
  if (parts) {
    const b = books.findIndex((x) => upper(x.name) === upper(parts.book))
    if (b < 0) return -1
    if (parts.sheet === null) return externalSheetIndex(b, BOOK_SCOPE)
    const s = books[b].sheets.findIndex((x) => upper(x.name) === upper(parts.sheet!))
    return s < 0 || s >= BOOK_SCOPE ? -1 : externalSheetIndex(b, s)
  }
  // 'Price List.xlsx'!Rate: a book's name where a sheet's would be (a sheet of this workbook wins).
  if (sheetNames.some((s) => upper(s) === upper(name))) return -1
  const b = books.findIndex((x) => upper(x.name) === upper(name))
  return b < 0 ? -1 : externalSheetIndex(b, BOOK_SCOPE)
}
