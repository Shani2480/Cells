import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { Axis } from "./axis"
import type { CellData } from "./cell-store"
import { moveMapper, rewriteFormula, shiftMapper, type RefMapper } from "./formula/adjust"
import type { Sheet } from "./sheet"

/** One cell's new content; `data: undefined` clears it. */
export type CellWrite = { row: number; col: number; data: CellData | undefined }

export type AxisName = "rows" | "cols"

/** One row's or column's size and visibility; `size: null` means the default size. */
export type AxisEntry = { index: number; size: number | null; hidden: boolean }

/**
 * Every change to a workbook is one of these: plain, serialisable data. `applyOp` performs one and returns the
 * operations that undo it, so undo/redo, autosave replay and the worker protocol all share one representation.
 */
export type Op =
  | { kind: "setCells"; sheet: number; cells: CellWrite[] }
  | { kind: "insertRows"; sheet: number; at: number; count: number }
  | { kind: "deleteRows"; sheet: number; at: number; count: number }
  | { kind: "insertCols"; sheet: number; at: number; count: number }
  | { kind: "deleteCols"; sheet: number; at: number; count: number }
  /** Insert blank cells in `range`, pushing the cells there down or right (within the range's columns or rows). */
  | { kind: "insertCells"; sheet: number; range: Range; shift: "down" | "right" }
  /** Delete the cells in `range`, pulling the cells below up or the cells to the right left. */
  | { kind: "deleteCells"; sheet: number; range: Range; shift: "up" | "left" }
  | { kind: "axis"; sheet: number; axis: AxisName; entries: AxisEntry[] }
  /**
   * Cut and paste: the cells in `source` move by (dr, dc), overwriting what was there. References anywhere in the
   * workbook that point into the source follow the cells.
   */
  | { kind: "moveCells"; sheet: number; source: Range; dr: number; dc: number }

/** What an operation touched, for redraws and (from Phase 2) recalculation. */
export type Change = { sheet: number; structural: boolean; range?: Range }

export type Applied = { inverse: Op[]; change: Change }

export function applyOp(sheets: readonly Sheet[], op: Op): Applied {
  const sheet = sheets[op.sheet]
  if (!sheet) {
    throw new Error(`No sheet ${op.sheet}`)
  }
  switch (op.kind) {
    case "setCells":
      return setCells(sheet, op)
    case "insertRows":
      return insertLines(sheets, op.sheet, "rows", op.at, op.count)
    case "deleteRows":
      return deleteLines(sheets, op.sheet, "rows", op.at, op.count)
    case "insertCols":
      return insertLines(sheets, op.sheet, "cols", op.at, op.count)
    case "deleteCols":
      return deleteLines(sheets, op.sheet, "cols", op.at, op.count)
    case "insertCells":
      return insertCells(sheets, op.sheet, op.range, op.shift)
    case "deleteCells":
      return deleteCells(sheets, op.sheet, op.range, op.shift)
    case "axis":
      return setAxis(sheet, op.sheet, op.axis, op.entries)
    case "moveCells":
      return moveCells(sheets, op)
  }
}

/**
 * Rewrites the references of every formula in the workbook that points at sheet `target`, before the cells move.
 * Returns setCells ops that put the old texts back; they run after the structural inverse, which has moved the
 * cells back to where they were when the texts were recorded.
 */
function rewriteReferences(sheets: readonly Sheet[], target: number, map: RefMapper): Op[] {
  const targetName = sheets[target].name
  const undo: Op[] = []
  sheets.forEach((sheet, index) => {
    const edits: { row: number; col: number; text: string }[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      // Formulas on other sheets can only point here through a sheet prefix.
      if (index !== target && !input.includes("!")) return
      const text = rewriteFormula(input, sheet.name, targetName, map)
      if (text !== null) edits.push({ row, col, text })
    })
    if (!edits.length) return
    const previous: CellWrite[] = []
    for (const edit of edits) {
      const data = sheet.cells.get(edit.row, edit.col)!
      previous.push({ row: edit.row, col: edit.col, data })
      sheet.cells.set(edit.row, edit.col, { ...data, input: edit.text })
    }
    undo.push({ kind: "setCells", sheet: index, cells: previous })
  })
  return undo
}

function setCells(sheet: Sheet, op: Extract<Op, { kind: "setCells" }>): Applied {
  const previous: CellWrite[] = []
  let r0 = MAX_ROWS
  let c0 = MAX_COLS
  let r1 = -1
  let c1 = -1
  for (const write of op.cells) {
    previous.push({ row: write.row, col: write.col, data: sheet.cells.get(write.row, write.col) })
    r0 = Math.min(r0, write.row)
    r1 = Math.max(r1, write.row)
    c0 = Math.min(c0, write.col)
    c1 = Math.max(c1, write.col)
  }
  for (const write of op.cells) {
    sheet.cells.set(write.row, write.col, write.data)
  }
  // Undo writes the old contents in reverse, so a cell written twice ends with its first old content.
  previous.reverse()
  return {
    inverse: [{ kind: "setCells", sheet: op.sheet, cells: previous }],
    change: { sheet: op.sheet, structural: false, range: r1 >= 0 ? { r0, c0, r1, c1 } : undefined },
  }
}

function axisOf(sheet: Sheet, axis: AxisName): Axis {
  return axis === "rows" ? sheet.rows : sheet.cols
}

function lineRange(axis: AxisName, from: number, to: number): Range {
  return axis === "rows" ? { r0: from, r1: to, c0: 0, c1: MAX_COLS - 1 } : { r0: 0, r1: MAX_ROWS - 1, c0: from, c1: to }
}

function insertLines(sheets: readonly Sheet[], index: number, axis: AxisName, at: number, count: number): Applied {
  const sheet = sheets[index]
  const size = axis === "rows" ? MAX_ROWS : MAX_COLS
  // Lines pushed off the end are lost; keep them so undo can restore them.
  const dropped = sheet.cells.entries(lineRange(axis, size - count, size - 1))
  const droppedAxis = axisOf(sheet, axis).overrides(size - count, size - 1)
  const texts = rewriteReferences(sheets, index, shiftMapper(axis, at, count))
  if (axis === "rows") {
    sheet.cells.shiftRows(at, count)
  } else {
    sheet.cells.shiftCols(at, count)
  }
  axisOf(sheet, axis).insert(at, count)
  const inverse: Op[] = [{ kind: axis === "rows" ? "deleteRows" : "deleteCols", sheet: index, at, count }]
  if (dropped.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: dropped })
  }
  if (droppedAxis.length) {
    inverse.push({ kind: "axis", sheet: index, axis, entries: droppedAxis })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function deleteLines(sheets: readonly Sheet[], index: number, axis: AxisName, at: number, count: number): Applied {
  const sheet = sheets[index]
  const range = lineRange(axis, at, at + count - 1)
  const removed = sheet.cells.entries(range)
  const removedAxis = axisOf(sheet, axis).overrides(at, at + count - 1)
  const texts = rewriteReferences(sheets, index, shiftMapper(axis, at, -count))
  sheet.cells.clear(range)
  if (axis === "rows") {
    sheet.cells.shiftRows(at + count, -count)
  } else {
    sheet.cells.shiftCols(at + count, -count)
  }
  axisOf(sheet, axis).remove(at, count)
  const inverse: Op[] = [{ kind: axis === "rows" ? "insertRows" : "insertCols", sheet: index, at, count }]
  if (removed.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: removed })
  }
  if (removedAxis.length) {
    inverse.push({ kind: "axis", sheet: index, axis, entries: removedAxis })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function insertCells(sheets: readonly Sheet[], index: number, range: Range, shift: "down" | "right"): Applied {
  const sheet = sheets[index]
  let dropped
  let texts: Op[]
  if (shift === "down") {
    const n = range.r1 - range.r0 + 1
    texts = rewriteReferences(sheets, index, shiftMapper("rows", range.r0, n, [range.c0, range.c1]))
    dropped = sheet.cells.entries({ r0: MAX_ROWS - n, r1: MAX_ROWS - 1, c0: range.c0, c1: range.c1 })
    sheet.cells.shiftRows(range.r0, n, range.c0, range.c1)
  } else {
    const n = range.c1 - range.c0 + 1
    texts = rewriteReferences(sheets, index, shiftMapper("cols", range.c0, n, [range.r0, range.r1]))
    dropped = sheet.cells.entries({ r0: range.r0, r1: range.r1, c0: MAX_COLS - n, c1: MAX_COLS - 1 })
    sheet.cells.shiftCols(range.c0, n, range.r0, range.r1)
  }
  const inverse: Op[] = [{ kind: "deleteCells", sheet: index, range, shift: shift === "down" ? "up" : "left" }]
  if (dropped.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: dropped })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function deleteCells(sheets: readonly Sheet[], index: number, range: Range, shift: "up" | "left"): Applied {
  const sheet = sheets[index]
  const removed = sheet.cells.entries(range)
  const texts =
    shift === "up"
      ? rewriteReferences(
          sheets,
          index,
          shiftMapper("rows", range.r0, -(range.r1 - range.r0 + 1), [range.c0, range.c1]),
        )
      : rewriteReferences(
          sheets,
          index,
          shiftMapper("cols", range.c0, -(range.c1 - range.c0 + 1), [range.r0, range.r1]),
        )
  sheet.cells.clear(range)
  if (shift === "up") {
    sheet.cells.shiftRows(range.r1 + 1, -(range.r1 - range.r0 + 1), range.c0, range.c1)
  } else {
    sheet.cells.shiftCols(range.c1 + 1, -(range.c1 - range.c0 + 1), range.r0, range.r1)
  }
  const inverse: Op[] = [{ kind: "insertCells", sheet: index, range, shift: shift === "up" ? "down" : "right" }]
  if (removed.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: removed })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function setAxis(sheet: Sheet, index: number, name: AxisName, entries: AxisEntry[]): Applied {
  const axis = axisOf(sheet, name)
  const previous: AxisEntry[] = entries.map((entry) => ({
    index: entry.index,
    size: axis.hasCustomSize(entry.index) ? axis.storedSize(entry.index) : null,
    hidden: axis.isHidden(entry.index),
  }))
  for (const entry of entries) {
    axis.setSize(entry.index, entry.size)
    axis.setHidden(entry.index, entry.hidden)
  }
  previous.reverse()
  return {
    inverse: [{ kind: "axis", sheet: index, axis: name, entries: previous }],
    change: { sheet: index, structural: true },
  }
}

function moveCells(sheets: readonly Sheet[], op: Extract<Op, { kind: "moveCells" }>): Applied {
  const { source, dr, dc } = op
  const sheet = sheets[op.sheet]
  // The texts' undo goes last: it restores formulas at their old places once the cells are back.
  const texts = rewriteReferences(sheets, op.sheet, moveMapper(source, dr, dc))
  const moved = sheet.cells.entries(source)
  const dest = { r0: source.r0 + dr, r1: source.r1 + dr, c0: source.c0 + dc, c1: source.c1 + dc }
  const previous = new Map<string, CellWrite>()
  const remember = (row: number, col: number) => {
    const key = `${row},${col}`
    if (!previous.has(key)) previous.set(key, { row, col, data: sheet.cells.get(row, col) })
  }
  for (const entry of moved) remember(entry.row, entry.col)
  for (const entry of sheet.cells.entries(dest)) remember(entry.row, entry.col)
  for (const entry of moved) remember(entry.row + dr, entry.col + dc)
  sheet.cells.clear(source)
  sheet.cells.clear(dest)
  for (const entry of moved) {
    const data = { ...entry.data }
    delete data.spill
    sheet.cells.set(entry.row + dr, entry.col + dc, data)
  }
  const inverse: Op[] = [{ kind: "setCells", sheet: op.sheet, cells: [...previous.values()] }, ...texts]
  return { inverse, change: { sheet: op.sheet, structural: true } }
}
