import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { Sheet } from "./sheet"
import { EXPLICIT_DEFAULT_ID, patchKey, type StylePatch, type StyleTable } from "./styles"

/**
 * Formatting on the sheet's three layers (cells, whole rows, whole columns), the way Excel stores it: formatting all
 * of column B sets the column's format plus the cells already stored in it, so a million empty cells cost nothing.
 */

/** Exact style ids to put back (undo of a format), flat: cells as [row, col, id, …], lines as [index, id, …]. */
export type StyleRecord = { cells: number[]; rows: number[]; cols: number[] }

const wholeColumns = (r: Range) => r.r0 === 0 && r.r1 === MAX_ROWS - 1
const wholeRows = (r: Range) => r.c0 === 0 && r.c1 === MAX_COLS - 1

/** Applies `patch` to every cell of `ranges`. Returns what undoes it. */
export function formatRanges(
  sheet: Sheet,
  styles: StyleTable,
  ranges: readonly Range[],
  patch: StylePatch,
): StyleRecord {
  const key = patchKey(patch)
  const undo: StyleRecord = { cells: [], rows: [], cols: [] }
  const seen = new Set<number>()
  const derive = (base: number, inherited: boolean) => {
    const id = styles.deriveCached(base, patch, key)
    // Back to the default inside a formatted row or column: keep it, rather than falling back to the line's format.
    return id === 0 && inherited ? EXPLICIT_DEFAULT_ID : id
  }
  const setCell = (row: number, col: number) => {
    const k = row * MAX_COLS + col
    if (seen.has(k)) return
    seen.add(k)
    const own = sheet.cells.style(row, col)
    const lineStyle = sheet.rowStyles.get(row) || sheet.colStyles.get(col)
    const base = own || lineStyle
    const next = derive(base, lineStyle !== 0)
    if (next === own) return
    undo.cells.push(row, col, own)
    sheet.cells.setStyle(row, col, next === EXPLICIT_DEFAULT_ID && lineStyle === 0 ? 0 : next)
  }
  for (const range of ranges) {
    if (wholeColumns(range) && wholeRows(range)) {
      // The whole sheet: every column's format, then the rows and cells that have their own.
      formatLines(sheet.colStyles, range.c0, range.c1, derive, undo.cols)
      formatLines(sheet.rowStyles, 0, MAX_ROWS - 1, derive, undo.rows, true)
      storedCells(sheet, range, setCell)
    } else if (wholeColumns(range)) {
      formatLines(sheet.colStyles, range.c0, range.c1, derive, undo.cols)
      // Formatted rows meet the columns in cells that take both formats, as in Excel.
      for (const [row] of sheet.rowStyles.entries()) for (let c = range.c0; c <= range.c1; c++) setCell(row, c)
      storedCells(sheet, range, setCell)
    } else if (wholeRows(range)) {
      formatLines(sheet.rowStyles, range.r0, range.r1, derive, undo.rows)
      for (const [col] of sheet.colStyles.entries()) for (let r = range.r0; r <= range.r1; r++) setCell(r, col)
      storedCells(sheet, range, setCell)
    } else {
      for (let c = range.c0; c <= range.c1; c++) for (let r = range.r0; r <= range.r1; r++) setCell(r, c)
    }
  }
  return undo
}

function formatLines(
  lines: Sheet["rowStyles"],
  from: number,
  to: number,
  derive: (base: number, inherited: boolean) => number,
  undo: number[],
  onlyStyled = false,
) {
  const visit = (i: number) => {
    const before = lines.get(i)
    const next = derive(before, false)
    if (next === before) return
    undo.push(i, before)
    lines.set(i, next)
  }
  if (onlyStyled) for (const [i] of lines.entries(from, to)) visit(i)
  else for (let i = from; i <= to; i++) visit(i)
}

function storedCells(sheet: Sheet, range: Range, visit: (row: number, col: number) => void) {
  const cells: number[] = []
  sheet.cells.forEachStyle(range, (row, col) => cells.push(row, col))
  for (let i = 0; i < cells.length; i += 2) visit(cells[i], cells[i + 1])
}

/** Puts exact style ids back. Returns the record that undoes this in turn. */
export function restoreStyles(sheet: Sheet, record: StyleRecord): StyleRecord {
  const undo: StyleRecord = { cells: [], rows: [], cols: [] }
  // Backwards: an id recorded twice ends with its first (oldest) value.
  for (let i = record.cols.length - 2; i >= 0; i -= 2) {
    undo.cols.push(record.cols[i], sheet.colStyles.get(record.cols[i]))
    sheet.colStyles.set(record.cols[i], record.cols[i + 1])
  }
  for (let i = record.rows.length - 2; i >= 0; i -= 2) {
    undo.rows.push(record.rows[i], sheet.rowStyles.get(record.rows[i]))
    sheet.rowStyles.set(record.rows[i], record.rows[i + 1])
  }
  for (let i = record.cells.length - 3; i >= 0; i -= 3) {
    const [row, col, id] = [record.cells[i], record.cells[i + 1], record.cells[i + 2]]
    undo.cells.push(row, col, sheet.cells.style(row, col))
    sheet.cells.setStyle(row, col, id)
  }
  return undo
}

/** The bounding box a style record touches (for redraws), or undefined for nothing. */
export function recordBounds(record: StyleRecord): Range | undefined {
  if (record.rows.length || record.cols.length) {
    let r0 = MAX_ROWS
    let r1 = -1
    let c0 = MAX_COLS
    let c1 = -1
    for (let i = 0; i < record.rows.length; i += 2) {
      r0 = Math.min(r0, record.rows[i])
      r1 = Math.max(r1, record.rows[i])
    }
    for (let i = 0; i < record.cols.length; i += 2) {
      c0 = Math.min(c0, record.cols[i])
      c1 = Math.max(c1, record.cols[i])
    }
    if (record.cols.length && !record.rows.length) return { r0: 0, r1: MAX_ROWS - 1, c0, c1 }
    if (record.rows.length && !record.cols.length) return { r0, r1, c0: 0, c1: MAX_COLS - 1 }
    return { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
  }
  if (!record.cells.length) return undefined
  let r0 = MAX_ROWS
  let r1 = -1
  let c0 = MAX_COLS
  let c1 = -1
  for (let i = 0; i < record.cells.length; i += 3) {
    r0 = Math.min(r0, record.cells[i])
    r1 = Math.max(r1, record.cells[i])
    c0 = Math.min(c0, record.cells[i + 1])
    c1 = Math.max(c1, record.cells[i + 1])
  }
  return { r0, r1, c0, c1 }
}

/** Shifts merged ranges for inserted (count > 0) or deleted (count < 0) rows or columns at `at`. */
export function shiftMerges(sheet: Sheet, axis: "rows" | "cols", at: number, count: number) {
  if (!sheet.merges.size) return
  const lo = axis === "rows" ? "r0" : "c0"
  const hi = axis === "rows" ? "r1" : "c1"
  const limit = (axis === "rows" ? MAX_ROWS : MAX_COLS) - 1
  const out: Range[] = []
  for (const m of sheet.merges.all) {
    const next = { ...m }
    if (count > 0) {
      // Inserting inside a merge widens it; before it, moves it.
      if (m[lo] >= at) next[lo] += count
      if (m[hi] >= at) next[hi] += count
      if (next[lo] > limit) continue
      next[hi] = Math.min(next[hi], limit)
    } else {
      const end = at - count - 1
      const cut = (v: number) => (v < at ? v : v > end ? v + count : at)
      next[lo] = cut(m[lo])
      // A deleted end pulls back to the last line that stays.
      next[hi] = m[hi] > end ? m[hi] + count : m[hi] >= at ? at - 1 : m[hi]
      if (next[hi] < next[lo]) continue
    }
    if (next.r0 === next.r1 && next.c0 === next.c1) continue
    out.push(next)
  }
  sheet.merges.replace(out)
}
