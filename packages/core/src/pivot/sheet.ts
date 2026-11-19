import { MAX_COLS, MAX_ROWS, type Range } from "../address"
import type { Op } from "../ops"
import type { Sheet } from "../sheet"
import type { PivotCache, PivotTable } from "./model"

/**
 * PivotTables on a sheet: copies for undo, where they sit, how they move with inserted and deleted rows and columns,
 * and which edits they refuse (Excel: "We can't make this change for the selected cells because it will affect a
 * PivotTable").
 */

export const PIVOT_MESSAGE =
  "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to " +
  "change the report. If you are trying to insert or delete cells, move the PivotTable and try again."

export function copyPivotTable(t: PivotTable): PivotTable {
  return {
    ...t,
    at: { ...t.at },
    location: { ...t.location },
    fields: t.fields.map((f) => ({ ...f, items: f.items.map((i) => ({ ...i })), subtotals: [...f.subtotals] })),
    rows: [...t.rows],
    cols: [...t.cols],
    pages: t.pages.map((p) => ({ ...p })),
    values: t.values.map((v) => ({ ...v })),
    captions: { ...t.captions },
    style: { ...t.style },
  }
}

export const copyPivotTables = (list: readonly PivotTable[]): PivotTable[] => list.map(copyPivotTable)

/** A cache's copy shares its records (they are only ever replaced whole, by a refresh). */
export function copyPivotCache(c: PivotCache): PivotCache {
  return {
    ...c,
    source: { ...c.source },
    fields: c.fields.map((f) => ({ ...f, items: f.items ? [...f.items] : null })),
  }
}

/** The cells a PivotTable takes: its table and the report filters above it (with the blank row between). */
export function pivotArea(t: PivotTable): Range {
  const loc = t.location
  if (!t.pages.length) return { ...loc }
  return { r0: Math.max(0, loc.r0 - t.pages.length - 1), c0: loc.c0, r1: loc.r1, c1: Math.max(loc.c1, loc.c0 + 1) }
}

const overlaps = (a: Range, b: Range) => a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/** The PivotTable at a cell (its filters included), if any. */
export function pivotAt(sheet: Sheet, row: number, col: number): PivotTable | null {
  for (const t of sheet.pivots) {
    const a = pivotArea(t)
    if (row >= a.r0 && row <= a.r1 && col >= a.c0 && col <= a.c1) return t
  }
  return null
}

/** PivotTables after inserting (count > 0) or deleting (count < 0) lines: moved, or gone when all their lines go. */
export function shiftPivots(list: readonly PivotTable[], axis: "rows" | "cols", at: number, count: number) {
  const out: PivotTable[] = []
  const max = axis === "rows" ? MAX_ROWS : MAX_COLS
  for (const t of list) {
    const a = pivotArea(t)
    const [lo, hi] = axis === "rows" ? [a.r0, a.r1] : [a.c0, a.c1]
    let by = 0
    if (count > 0) by = at <= lo ? count : 0
    else {
      const end = at - count - 1
      if (at <= lo && end >= hi) continue
      if (end < lo) by = count
    }
    if (!by) {
      out.push(t)
      continue
    }
    if (hi + by >= max) continue
    const moved = copyPivotTable(t)
    if (axis === "rows") {
      moved.at.row += by
      moved.location.r0 += by
      moved.location.r1 += by
    } else {
      moved.at.col += by
      moved.location.c0 += by
      moved.location.c1 += by
    }
    out.push(moved)
  }
  return out
}

/**
 * Whether an edit would change part of a PivotTable: writing into its cells, or inserting, deleting or moving cells
 * through it. Formats are fine (as in Excel); deleting all of a PivotTable's rows or columns deletes it.
 */
export function breaksPivot(sheet: Sheet, op: Op): boolean {
  if (!("sheet" in op) || !sheet.pivots.length) return false
  const areas = sheet.pivots.map(pivotArea)
  switch (op.kind) {
    case "setCells":
      return op.cells.some((w) => areas.some((a) => w.row >= a.r0 && w.row <= a.r1 && w.col >= a.c0 && w.col <= a.c1))
    case "insertRows":
    case "deleteRows":
    case "insertCols":
    case "deleteCols": {
      const rows = op.kind === "insertRows" || op.kind === "deleteRows"
      const deleting = op.kind === "deleteRows" || op.kind === "deleteCols"
      return areas.some((a) => {
        const [lo, hi] = rows ? [a.r0, a.r1] : [a.c0, a.c1]
        if (!deleting) return op.at > lo && op.at <= hi
        const end = op.at + op.count - 1
        return op.at <= hi && end >= lo && !(op.at <= lo && end >= hi)
      })
    }
    case "insertCells":
    case "deleteCells":
      return areas.some((a) => sameBand(a, op))
    case "moveCells": {
      const target = {
        r0: op.source.r0 + op.dr,
        c0: op.source.c0 + op.dc,
        r1: op.source.r1 + op.dr,
        c1: op.source.c1 + op.dc,
      }
      return areas.some((a) => overlaps(a, op.source) || overlaps(a, target))
    }
    default:
      return false
  }
}

/** Inserted or deleted cells shift the cells right of or below them, which moves part of a PivotTable there. */
function sameBand(a: Range, op: Extract<Op, { kind: "insertCells" | "deleteCells" }>): boolean {
  const r = op.range
  if (op.shift === "right" || op.shift === "left") return r.r0 <= a.r1 && r.r1 >= a.r0 && r.c0 <= a.c1
  return r.c0 <= a.c1 && r.c1 >= a.c0 && r.r0 <= a.r1
}
