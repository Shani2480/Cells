import type { CellValue } from "./values"
import type { Op } from "./ops"
import { Workbook } from "./workbook"

/**
 * Data › What-If Analysis. Trial values are worked out in a private copy of the workbook (its own calculation),
 * so the sheet, its undo history and the calculation worker only see the answer.
 */

export type CellRef = { sheet: number; row: number; col: number }

/** A copy of the workbook's cells and names that calculates on its own. */
export function calcCopy(wb: Workbook<unknown>): Workbook<unknown> {
  const copy = new Workbook(wb.sheets.map((s) => s.name))
  const ops: Op[] = []
  wb.sheets.forEach((sheet, index) => {
    const used = sheet.cells.usedBounds()
    if (used)
      ops.push({
        kind: "setCells",
        sheet: index,
        cells: sheet.cells.entries(used).map((e) => ({ row: e.row, col: e.col, data: { ...e.data } })),
      })
  })
  const names = wb.names.all()
  if (names.length) ops.push({ kind: "names", names: names.map((n) => ({ ...n })) })
  if (ops.length) copy.transact(ops, { label: "Copy", history: false })
  copy.calculate(true)
  return copy
}

function setNumber(wb: Workbook<unknown>, at: CellRef, value: number) {
  wb.transact([{ kind: "setCells", sheet: at.sheet, cells: [{ row: at.row, col: at.col, data: { value } }] }], {
    label: "Trial",
    history: false,
  })
}

const numberOf = (v: CellValue | undefined): number | null => (typeof v === "number" && Number.isFinite(v) ? v : null)

export type GoalSeekResult = {
  /** The changing cell's value found (or the closest tried). */
  value: number
  /** The set cell's value with it. */
  result: number | null
  found: boolean
  iterations: number
}

/**
 * Goal Seek: the value of `changing` that makes the formula in `set` equal `target`, as Excel does it (up to 100
 * tries, close enough within 0.001). Secant steps, with halving between two tries on either side of the target
 * once there are two.
 */
export function goalSeek(
  wb: Workbook<unknown>,
  set: CellRef,
  target: number,
  changing: CellRef,
  { maxIterations = 100, maxChange = 0.001 } = {},
): GoalSeekResult {
  const copy = calcCopy(wb)
  const f = (x: number): number | null => {
    setNumber(copy, changing, x)
    const v = numberOf(copy.sheets[set.sheet].cells.value(set.row, set.col))
    return v === null ? null : v - target
  }
  let x0 = numberOf(wb.sheets[changing.sheet].cells.value(changing.row, changing.col)) ?? 0
  let f0 = f(x0)
  let best = { x: x0, fx: f0 }
  const better = (x: number, fx: number | null) => {
    if (fx !== null && (best.fx === null || Math.abs(fx) < Math.abs(best.fx))) best = { x, fx }
  }
  if (f0 !== null && Math.abs(f0) <= maxChange) return done(best, true, 0)
  let x1 = x0 === 0 ? 1 : x0 * 1.01
  let f1 = f(x1)
  better(x1, f1)
  // Two tries on either side of the target, once found.
  let low: { x: number; fx: number } | null = null
  let high: { x: number; fx: number } | null = null
  const bracket = (x: number, fx: number | null) => {
    if (fx === null) return
    if (fx < 0) low = { x, fx }
    else high = { x, fx }
  }
  bracket(x0, f0)
  bracket(x1, f1)
  for (let i = 1; i <= maxIterations; i++) {
    if (f1 !== null && Math.abs(f1) <= maxChange) return done({ x: x1, fx: f1 }, true, i)
    let next: number
    if (low && high) {
      const l = low as { x: number; fx: number }
      const h = high as { x: number; fx: number }
      // Secant between the two sides (false position), halving when it stalls.
      next = l.x - (l.fx * (h.x - l.x)) / (h.fx - l.fx)
      if (!Number.isFinite(next) || i % 4 === 0) next = (l.x + h.x) / 2
    } else if (f0 !== null && f1 !== null && f1 !== f0) {
      next = x1 - (f1 * (x1 - x0)) / (f1 - f0)
      // Don't leap too far in one try.
      const step = next - x1
      const limit = Math.max(1, Math.abs(x1)) * 100
      if (Math.abs(step) > limit) next = x1 + Math.sign(step) * limit
    } else {
      next = x1 + (x1 - x0 || 1) * 2
    }
    if (!Number.isFinite(next)) break
    x0 = x1
    f0 = f1
    x1 = next
    f1 = f(x1)
    better(x1, f1)
    bracket(x1, f1)
  }
  return done(best, best.fx !== null && Math.abs(best.fx) <= maxChange, maxIterations)

  function done(b: { x: number; fx: number | null }, found: boolean, iterations: number): GoalSeekResult {
    return { value: b.x, result: b.fx === null ? null : b.fx + target, found, iterations }
  }
}
