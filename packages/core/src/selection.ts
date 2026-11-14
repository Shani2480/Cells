import { MAX_COLS, MAX_ROWS, rangeContains, rangeOf, type CellAddress, type Range } from "./address"

/**
 * A selection as Excel keeps it: one or more ranges (Ctrl+click adds), the active cell (white, where typing goes),
 * and the last range's fixed corner (`anchor`) and moving corner (`focus`) for Shift+click and Shift+arrows.
 */
export type Selection = {
  readonly ranges: readonly Range[]
  readonly active: CellAddress
  /** Index of the range that holds the active cell. */
  readonly activeRange: number
  readonly anchor: CellAddress
  readonly focus: CellAddress
}

export function selectCell(address: CellAddress): Selection {
  return { ranges: [rangeOf(address)], active: address, activeRange: 0, anchor: address, focus: address }
}

/** A single range with the active cell at `active` (defaults to the range's top-left). */
export function selectRange(range: Range, active?: CellAddress): Selection {
  const start = active ?? { row: range.r0, col: range.c0 }
  return {
    ranges: [range],
    active: start,
    activeRange: 0,
    anchor: { row: range.r0, col: range.c0 },
    focus: { row: range.r1, col: range.c1 },
  }
}

/** Shift+click / Shift+arrow: the last range spans anchor..focus; the active cell stays. */
export function extendSelection(selection: Selection, focus: CellAddress): Selection {
  const ranges = selection.ranges.slice(0, -1)
  ranges.push(rangeOf(selection.anchor, focus))
  const activeRange = ranges.length - 1
  const last = ranges[activeRange]
  const active = rangeContains(last, selection.active.row, selection.active.col) ? selection.active : selection.anchor
  return { ranges, active, activeRange, anchor: selection.anchor, focus }
}

/** Ctrl+click: a new range starting at `address`, keeping the others. */
export function addRange(selection: Selection, address: CellAddress, range: Range = rangeOf(address)): Selection {
  const ranges = [...selection.ranges, range]
  return {
    ranges,
    active: address,
    activeRange: ranges.length - 1,
    anchor: { row: range.r0, col: range.c0 },
    focus: { row: range.r1, col: range.c1 },
  }
}

/** Whole rows r0..r1 (clicking or dragging row headers); the active cell goes to `activeCol` of the first row. */
export function selectRows(r0: number, r1: number, activeCol: number, anchorRow = r0): Selection {
  const focusRow = anchorRow === r0 ? r1 : r0
  return {
    ranges: [{ r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: 0, c1: MAX_COLS - 1 }],
    active: { row: anchorRow, col: activeCol },
    activeRange: 0,
    anchor: { row: anchorRow, col: 0 },
    focus: { row: focusRow, col: MAX_COLS - 1 },
  }
}

export function selectCols(c0: number, c1: number, activeRow: number, anchorCol = c0): Selection {
  const focusCol = anchorCol === c0 ? c1 : c0
  return {
    ranges: [{ r0: 0, r1: MAX_ROWS - 1, c0: Math.min(c0, c1), c1: Math.max(c0, c1) }],
    active: { row: activeRow, col: anchorCol },
    activeRange: 0,
    anchor: { row: 0, col: anchorCol },
    focus: { row: MAX_ROWS - 1, col: focusCol },
  }
}

export function selectAll(active: CellAddress): Selection {
  return {
    ranges: [{ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }],
    active,
    activeRange: 0,
    anchor: { row: 0, col: 0 },
    focus: { row: MAX_ROWS - 1, col: MAX_COLS - 1 },
  }
}

/** Ctrl+Space: the selected ranges' columns. */
export function expandToColumns(selection: Selection): Selection {
  return expandAll(selection, (range) => ({ ...range, r0: 0, r1: MAX_ROWS - 1 }), "cols")
}

/** Shift+Space: the selected ranges' rows. */
export function expandToRows(selection: Selection): Selection {
  return expandAll(selection, (range) => ({ ...range, c0: 0, c1: MAX_COLS - 1 }), "rows")
}

function expandAll(selection: Selection, grow: (range: Range) => Range, kind: "rows" | "cols"): Selection {
  const ranges = selection.ranges.map(grow)
  const last = ranges[ranges.length - 1]
  const anchor = kind === "cols" ? { row: 0, col: selection.anchor.col } : { row: selection.anchor.row, col: 0 }
  const focus =
    kind === "cols" ? { row: MAX_ROWS - 1, col: selection.focus.col } : { row: selection.focus.row, col: MAX_COLS - 1 }
  return {
    ranges,
    active: selection.active,
    activeRange: selection.activeRange,
    anchor: kind === "cols" ? { ...anchor, col: clampTo(anchor.col, last.c0, last.c1) } : anchor,
    focus,
  }
}

function clampTo(value: number, low: number, high: number) {
  return Math.min(high, Math.max(low, value))
}

/** True when the selection is a single cell (Enter and Tab then move instead of cycling inside it). */
export function isSingleCell(selection: Selection) {
  if (selection.ranges.length !== 1) {
    return false
  }
  const range = selection.ranges[0]
  return range.r0 === range.r1 && range.c0 === range.c1
}

export function selectionContains(selection: Selection, row: number, col: number) {
  return selection.ranges.some((range) => rangeContains(range, row, col))
}

/** The range that encloses every selected range. */
export function selectionBounds(selection: Selection): Range {
  let { r0, c0, r1, c1 } = selection.ranges[0]
  for (const range of selection.ranges) {
    r0 = Math.min(r0, range.r0)
    c0 = Math.min(c0, range.c0)
    r1 = Math.max(r1, range.r1)
    c1 = Math.max(c1, range.c1)
  }
  return { r0, c0, r1, c1 }
}
