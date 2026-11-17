import type { CellValue } from "./values"
import type { CellWrite, Op } from "./ops"
import { columnName, type Range } from "./address"
import type { Sheet } from "./sheet"
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

// ── Data tables ─────────────────────────────────────────────────────────────────────────────────────────────────

/** The input of a What-If data table cell: =TABLE(row input, column input). */
export const isDataTableInput = (input: string | undefined): input is string => !!input && /^=TABLE\(/i.test(input)

/** The block of data table cells holding the same TABLE formula as the cell at row, col (or null if it holds none). */
export function dataTableBlockAt(sheet: Sheet, row: number, col: number): Range | null {
  const text = sheet.cells.input(row, col)
  if (!isDataTableInput(text)) return null
  const same = (r: number, c: number) => r >= 0 && c >= 0 && sheet.cells.input(r, c) === text
  let r0 = row
  while (same(r0 - 1, col)) r0--
  let c0 = col
  while (same(row, c0 - 1)) c0--
  let r1 = row
  while (same(r1 + 1, col)) r1++
  let c1 = col
  while (same(row, c1 + 1)) c1++
  return { r0, c0, r1, c1 }
}

/**
 * Whether an operation would change part of a data table ("Cannot change part of a data table."): writing some but
 * not all of its cells, or inserting, deleting or moving cells through it. Writing or clearing a whole table is fine.
 */
export function breaksDataTable(sheet: Sheet, op: Op): boolean {
  if (!("sheet" in op)) return false
  switch (op.kind) {
    case "setCells": {
      let written: Set<string> | null = null
      for (const w of op.cells) {
        if (!isDataTableInput(sheet.cells.input(w.row, w.col))) continue
        // A rewrite of the same TABLE formula (the Data Table dialog over an existing table) is fine.
        if (w.data?.input === sheet.cells.input(w.row, w.col)) continue
        const block = dataTableBlockAt(sheet, w.row, w.col)!
        written ??= new Set(op.cells.map((c) => `${c.row},${c.col}`))
        for (let r = block.r0; r <= block.r1; r++)
          for (let c = block.c0; c <= block.c1; c++) if (!written.has(`${r},${c}`)) return true
      }
      return false
    }
    case "insertRows":
    case "deleteRows":
    case "insertCols":
    case "deleteCols": {
      const rows = op.kind === "insertRows" || op.kind === "deleteRows"
      const deleting = op.kind === "deleteRows" || op.kind === "deleteCols"
      return dataTableBlocks(sheet).some((b) => {
        const [lo, hi] = rows ? [b.r0, b.r1] : [b.c0, b.c1]
        // Inserting inside a table splits it; deleting part of it (but not all) breaks it.
        if (!deleting) return op.at > lo && op.at <= hi
        const end = op.at + op.count - 1
        return op.at <= hi && end >= lo && !(op.at <= lo && end >= hi)
      })
    }
    case "insertCells":
    case "deleteCells":
    case "moveCells": {
      const area = op.kind === "moveCells" ? op.source : op.range
      return dataTableBlocks(sheet).some(
        (b) =>
          b.r0 <= area.r1 &&
          b.r1 >= area.r0 &&
          b.c0 <= area.c1 &&
          b.c1 >= area.c0 &&
          !(area.r0 <= b.r0 && area.r1 >= b.r1 && area.c0 <= b.c0 && area.c1 >= b.c1),
      )
    }
    default:
      return false
  }
}

/** Every data table block on the sheet. */
export function dataTableBlocks(sheet: Sheet): Range[] {
  const out: Range[] = []
  sheet.cells.forEachFormula((row, col, input) => {
    if (!isDataTableInput(input)) return
    if (out.some((b) => row >= b.r0 && row <= b.r1 && col >= b.c0 && col <= b.c1)) return
    out.push(dataTableBlockAt(sheet, row, col)!)
  })
  return out
}

/**
 * Data › What-If Analysis › Data Table: the TABLE formula for the cells of `range` past its first row and column
 * (the first row and column hold the formulas and the values to try). Null when the range is too small.
 */
export function dataTableOp(
  sheet: Sheet,
  sheetIndex: number,
  range: Range,
  rowInput: string | null,
  colInput: string | null,
): Op | null {
  if (range.r1 <= range.r0 || range.c1 <= range.c0 || (!rowInput && !colInput)) return null
  const input = `=TABLE(${rowInput ?? ""},${colInput ?? ""})`
  const cells: CellWrite[] = []
  for (let r = range.r0 + 1; r <= range.r1; r++)
    for (let c = range.c0 + 1; c <= range.c1; c++) {
      const style = sheet.cells.style(r, c)
      cells.push({ row: r, col: c, data: { input, value: 0, ...(style ? { style } : {}) } })
    }
  return { kind: "setCells", sheet: sheetIndex, cells }
}

// ── Scenario Summary ────────────────────────────────────────────────────────────────────────────────────────────

export type ScenarioSummary = {
  /** Scenario names, in order (the columns after Current Values). */
  names: string[]
  /** One row per changing cell: its reference and its value now and in each scenario. */
  changing: { ref: string; current: CellValue | undefined; values: (CellValue | undefined)[] }[]
  /** One row per result cell: its reference and its value now and with each scenario shown. */
  results: { ref: string; current: CellValue | undefined; values: (CellValue | undefined)[] }[]
}

/**
 * Scenario Manager › Summary: the changing cells of every scenario and the result cells' values with each scenario
 * shown, worked out in a private copy so the sheet stays as it is.
 */
export function scenarioSummary(wb: Workbook<unknown>, sheetIndex: number, resultCells: Range[]): ScenarioSummary {
  const sheet = wb.sheets[sheetIndex]
  const scenarios = sheet.scenarios.scenarios.filter((s) => !s.hidden)
  const key = (row: number, col: number) => `${row},${col}`
  const changingCells: { row: number; col: number }[] = []
  const seen = new Set<string>()
  for (const s of scenarios)
    for (const c of s.cells)
      if (!seen.has(key(c.row, c.col))) {
        seen.add(key(c.row, c.col))
        changingCells.push({ row: c.row, col: c.col })
      }
  const results: { row: number; col: number }[] = []
  for (const r of resultCells)
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) results.push({ row, col })
  const abs = (row: number, col: number) => `$${columnName(col)}$${row + 1}`
  const resultValues: (CellValue | undefined)[][] = results.map(() => [])
  for (const s of scenarios) {
    const copy = calcCopy(wb)
    const cells: CellWrite[] = s.cells.map((c) => ({ row: c.row, col: c.col, data: { value: c.value } }))
    copy.transact([{ kind: "setCells", sheet: sheetIndex, cells }], { label: "Trial", history: false })
    copy.calculate(true)
    results.forEach((r, i) => resultValues[i].push(copy.sheets[sheetIndex].cells.value(r.row, r.col)))
  }
  return {
    names: scenarios.map((s) => s.name),
    changing: changingCells.map(({ row, col }) => ({
      ref: abs(row, col),
      current: sheet.cells.value(row, col),
      values: scenarios.map((s) => s.cells.find((c) => c.row === row && c.col === col)?.value),
    })),
    results: results.map(({ row, col }, i) => ({
      ref: abs(row, col),
      current: sheet.cells.value(row, col),
      values: resultValues[i],
    })),
  }
}

/** The summary laid out as Excel's Scenario Summary sheet: B2 the title, the table from row 3, notes under it. */
export function scenarioSummaryCells(summary: ScenarioSummary): {
  cells: CellWrite[]
  /** Rows to show in bold (the headings) and the area of the table, for formats. */
  headings: Range[]
  table: Range
} {
  const cells: CellWrite[] = []
  const put = (row: number, col: number, value: CellValue | undefined) => {
    if (value !== undefined) cells.push({ row, col, data: { value } })
  }
  const lastCol = 3 + summary.names.length
  put(1, 1, "Scenario Summary")
  put(1, 3, "Current Values:")
  summary.names.forEach((n, i) => put(1, 4 + i, n))
  let row = 2
  put(row++, 1, "Changing Cells:")
  for (const c of summary.changing) {
    put(row, 2, c.ref)
    put(row, 3, c.current)
    c.values.forEach((v, i) => put(row, 4 + i, v))
    row++
  }
  const resultsRow = row
  if (summary.results.length) {
    put(row++, 1, "Result Cells:")
    for (const c of summary.results) {
      put(row, 2, c.ref)
      put(row, 3, c.current)
      c.values.forEach((v, i) => put(row, 4 + i, v))
      row++
    }
  }
  const table = { r0: 1, c0: 1, r1: row - 1, c1: lastCol }
  put(row++, 1, "Notes:  Current Values column represents values of changing cells at")
  put(row++, 1, "time Scenario Summary Report was created.  Changing cells for each")
  put(row++, 1, "scenario are highlighted in gray.")
  const headings = [
    { r0: 1, c0: 1, r1: 1, c1: lastCol },
    { r0: 2, c0: 1, r1: 2, c1: 1 },
  ]
  if (summary.results.length) headings.push({ r0: resultsRow, c0: 1, r1: resultsRow, c1: 1 })
  return { cells, headings, table }
}
