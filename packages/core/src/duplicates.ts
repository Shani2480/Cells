import type { Range } from "./address"
import type { CellData } from "./cell-store"
import { displayText } from "./format"
import type { CellWrite, Op } from "./ops"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"
import { isError } from "./values"

/**
 * Data › Remove Duplicates, as Excel's: a row of the range is a duplicate when every chosen column matches an earlier
 * row. Numbers match when both their value and the text they show are the same (46100 as a date isn't 46100 as a
 * number; =0.1+0.2 isn't 0.3), errors by code, and text and logicals by what they show, ignoring case. Blanks match.
 * Duplicate rows go from the range only: the rows that stay move up inside it; cells beside and below it stay.
 */
export type DuplicatesResult = { ops: Op[]; removed: number; remain: number }

function keyOf(data: CellData | undefined, code: string | undefined): string {
  const v = data?.value
  if (v === undefined || v === "") return "b"
  if (typeof v === "number") return `n${v}|${displayText(v, code)}`
  if (isError(v)) return `e${v.error}`
  return `t${displayText(v, code).toLowerCase()}`
}

export function removeDuplicatesOps(
  sheetIndex: number,
  sheet: Sheet,
  styles: StyleTable,
  range: Range,
  columns: readonly number[],
  header: boolean,
): DuplicatesResult {
  const first = range.r0 + (header ? 1 : 0)
  const seen = new Set<string>()
  const duplicates: number[] = []
  for (let row = first; row <= range.r1; row++) {
    const key = columns
      .map((col) => keyOf(sheet.cells.get(row, col), styles.get(sheet.styleAt(row, col)).numberFormat))
      .join("\u0000")
    if (seen.has(key)) duplicates.push(row)
    else seen.add(key)
  }
  if (!duplicates.length) return { ops: [], removed: 0, remain: range.r1 - first + 1 }
  // The rows that stay move up inside the range, with their formats; the rows freed at its foot are emptied.
  const gone = new Set(duplicates)
  const cells: CellWrite[] = []
  let to = first
  for (let row = first; row <= range.r1; row++) {
    if (gone.has(row)) continue
    if (row !== to)
      for (let col = range.c0; col <= range.c1; col++) {
        const data = sheet.cells.get(row, col)
        cells.push({ row: to, col, data: data && !data.spill ? { ...data } : undefined })
      }
    to++
  }
  for (let row = to; row <= range.r1; row++)
    for (let col = range.c0; col <= range.c1; col++) cells.push({ row, col, data: undefined })
  const ops: Op[] = [{ kind: "setCells", sheet: sheetIndex, cells }]
  return { ops, removed: duplicates.length, remain: range.r1 - first + 1 - duplicates.length }
}
