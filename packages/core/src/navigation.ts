import { MAX_COLS, MAX_ROWS, rangeOf, type CellAddress, type Range } from "./address"
import type { Axis } from "./axis"
import { isSingleCell, type Selection } from "./selection"
import type { Sheet } from "./sheet"

export type Direction = "up" | "down" | "left" | "right"

function isVertical(dir: Direction) {
  return dir === "up" || dir === "down"
}

function sign(dir: Direction): 1 | -1 {
  return dir === "down" || dir === "right" ? 1 : -1
}

/** The next shown index after `index` in `dir`, or -1 at the edge. */
function nextShown(axis: Axis, index: number, dir: 1 | -1): number {
  for (let i = index + dir; i >= 0 && i < axis.count; i += dir) {
    if (!axis.isHidden(i)) {
      return i
    }
  }
  return -1
}

/** The last shown index in `dir` (the sheet's edge as the arrow keys reach it). */
function edgeShown(axis: Axis, dir: 1 | -1): number {
  return axis.firstVisibleFrom(dir === 1 ? axis.count - 1 : 0, dir === 1 ? -1 : 1)
}

/** One cell in `dir`, skipping hidden rows/columns; stays put at the edge. */
export function step(sheet: Sheet, from: CellAddress, dir: Direction): CellAddress {
  const vertical = isVertical(dir)
  const axis = vertical ? sheet.rows : sheet.cols
  const next = nextShown(axis, vertical ? from.row : from.col, sign(dir))
  if (next === -1) {
    return from
  }
  return vertical ? { row: next, col: from.col } : { row: from.row, col: next }
}

function hasValueAt(sheet: Sheet, vertical: boolean, line: number, index: number) {
  return vertical ? sheet.cells.hasValue(index, line) : sheet.cells.hasValue(line, index)
}

/** The next index after `from` along the line that holds a value and is shown, or -1. */
function nextValue(sheet: Sheet, vertical: boolean, line: number, from: number, dir: 1 | -1): number {
  const axis = vertical ? sheet.rows : sheet.cols
  if (vertical) {
    let row = sheet.cells.nextValueRow(line, from, dir)
    while (row !== -1 && axis.isHidden(row)) {
      row = sheet.cells.nextValueRow(line, row, dir)
    }
    return row
  }
  const cols = sheet.cells.usedColumns()
  if (dir === 1) {
    for (const col of cols) {
      if (col > from && !axis.isHidden(col) && sheet.cells.hasValue(line, col)) return col
    }
  } else {
    for (let i = cols.length - 1; i >= 0; i--) {
      const col = cols[i]
      if (col < from && !axis.isHidden(col) && sheet.cells.hasValue(line, col)) return col
    }
  }
  return -1
}

/**
 * Ctrl+arrow (and End, arrow): from a filled cell whose neighbour is filled, to the last filled cell of that run;
 * otherwise to the next filled cell; with none, to the sheet's edge. Hidden rows and columns are passed over.
 */
export function dataEdge(sheet: Sheet, from: CellAddress, dir: Direction): CellAddress {
  const vertical = isVertical(dir)
  const axis = vertical ? sheet.rows : sheet.cols
  const d = sign(dir)
  const line = vertical ? from.col : from.row
  let index = vertical ? from.row : from.col
  const at = (i: number) => (vertical ? { row: i, col: line } : { row: line, col: i })

  const next = nextShown(axis, index, d)
  if (next === -1) {
    return from
  }
  if (hasValueAt(sheet, vertical, line, index) && hasValueAt(sheet, vertical, line, next)) {
    index = next
    for (;;) {
      const after = nextShown(axis, index, d)
      if (after === -1 || !hasValueAt(sheet, vertical, line, after)) {
        return at(index)
      }
      index = after
    }
  }
  const found = nextValue(sheet, vertical, line, index, d)
  return at(found === -1 ? edgeShown(axis, d) : found)
}

/** Ctrl+End: the bottom-right corner of the used range (or A1 on an empty sheet). */
export function lastCell(sheet: Sheet): CellAddress {
  const bounds = sheet.cells.usedBounds()
  return bounds ? { row: bounds.r1, col: bounds.c1 } : { row: 0, col: 0 }
}

/**
 * The current region (Ctrl+A in data, Ctrl+Shift+*): the block around `from` bounded by empty rows and columns,
 * growing through diagonal neighbours as Excel's does.
 */
export function currentRegion(sheet: Sheet, from: CellAddress): Range {
  const range = rangeOf(from)
  const filled = (r: number, c: number) =>
    r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS && sheet.cells.hasValue(r, c)
  const rowHas = (r: number, c0: number, c1: number) => {
    for (let c = Math.max(0, c0); c <= Math.min(MAX_COLS - 1, c1); c++) if (filled(r, c)) return true
    return false
  }
  const colHas = (c: number, r0: number, r1: number) => {
    // Columns are sparse; skip straight to the next value in the column.
    if (c < 0 || c >= MAX_COLS || !sheet.cells.columnHasValues(c)) return false
    const first = Math.max(0, r0)
    if (sheet.cells.hasValue(first, c)) return true
    const next = sheet.cells.nextValueRow(c, first, 1)
    return next !== -1 && next <= Math.min(MAX_ROWS - 1, r1)
  }
  for (let grew = true; grew;) {
    grew = false
    if (range.r0 > 0 && rowHas(range.r0 - 1, range.c0 - 1, range.c1 + 1)) {
      range.r0--
      grew = true
    }
    if (range.r1 < MAX_ROWS - 1 && rowHas(range.r1 + 1, range.c0 - 1, range.c1 + 1)) {
      // Grow down in one go to the end of the run in any column of the block, rather than a row at a time.
      let bottom = range.r1 + 1
      for (let c = range.c0; c <= range.c1; c++) {
        if (sheet.cells.hasValue(range.r1 + 1, c) || sheet.cells.hasValue(range.r1, c)) {
          const end = runEnd(sheet, c, range.r1 + 1)
          bottom = Math.max(bottom, end)
        }
      }
      range.r1 = bottom
      grew = true
    }
    if (range.c0 > 0 && colHas(range.c0 - 1, range.r0 - 1, range.r1 + 1)) {
      range.c0--
      grew = true
    }
    if (range.c1 < MAX_COLS - 1 && colHas(range.c1 + 1, range.r0 - 1, range.r1 + 1)) {
      range.c1++
      grew = true
    }
  }
  return range
}

/** The last row of the run of filled cells in `col` that contains or starts right after `row`. */
function runEnd(sheet: Sheet, col: number, row: number): number {
  let end = row
  while (end + 1 < MAX_ROWS && sheet.cells.hasValue(end + 1, col)) end++
  return sheet.cells.hasValue(row, col) ? end : row
}

/**
 * Enter / Tab (and their Shift forms) inside a multi-cell selection: the active cell walks the selection, row by
 * row for Tab and column by column for Enter, wrapping to the next range. With a single cell selected, it is a
 * plain step.
 */
export function moveWithinSelection(sheet: Sheet, selection: Selection, dir: Direction): Selection {
  if (isSingleCell(selection)) {
    const active = step(sheet, selection.active, dir)
    return { ranges: [rangeOf(active)], active, activeRange: 0, anchor: active, focus: active }
  }
  const vertical = isVertical(dir)
  const d = sign(dir)
  let rangeIndex = selection.activeRange
  let { row, col } = selection.active
  // Bounded walk: at most every cell of every range once, skipping hidden cells.
  for (let guard = 0; guard < 4_000_000; guard++) {
    const range = selection.ranges[rangeIndex]
    if (vertical) {
      row += d
      if (row > range.r1 || row < range.r0) {
        col += d
        if (col > range.c1 || col < range.c0) {
          rangeIndex = (rangeIndex + d + selection.ranges.length) % selection.ranges.length
          const nextRange = selection.ranges[rangeIndex]
          row = d === 1 ? nextRange.r0 : nextRange.r1
          col = d === 1 ? nextRange.c0 : nextRange.c1
        } else {
          row = d === 1 ? range.r0 : range.r1
        }
      }
    } else {
      col += d
      if (col > range.c1 || col < range.c0) {
        row += d
        if (row > range.r1 || row < range.r0) {
          rangeIndex = (rangeIndex + d + selection.ranges.length) % selection.ranges.length
          const nextRange = selection.ranges[rangeIndex]
          row = d === 1 ? nextRange.r0 : nextRange.r1
          col = d === 1 ? nextRange.c0 : nextRange.c1
        } else {
          col = d === 1 ? range.c0 : range.c1
        }
      }
    }
    if (!sheet.rows.isHidden(row) && !sheet.cols.isHidden(col)) {
      break
    }
  }
  return { ...selection, active: { row, col }, activeRange: rangeIndex }
}
