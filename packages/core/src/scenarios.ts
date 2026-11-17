import type { Range } from "./address"
import { shiftRange } from "./conditional"
import type { CellWrite, Op } from "./ops"
import type { Sheet } from "./sheet"
import type { CellValue } from "./values"

/**
 * Data › What-If Analysis › Scenario Manager: named sets of values for a sheet's changing cells. Kept per sheet, as
 * Excel does (<scenarios> in the sheet's XML).
 */

export type ScenarioCell = { row: number; col: number; value: CellValue }

export type Scenario = {
  name: string
  comment: string
  /** Who added it (Excel shows "Created by … on …" in the comment). */
  user: string
  /** Protection › Prevent changes / Hide: they apply when the sheet is protected. */
  locked: boolean
  hidden: boolean
  cells: ScenarioCell[]
}

export type ScenarioSet = {
  scenarios: Scenario[]
  /** The scenario shown last and the one selected in the list (indexes), as Excel keeps them. */
  current: number | null
  show: number | null
  /** The result cells chosen for the last summary. */
  resultCells: Range[]
}

export const emptyScenarios = (): ScenarioSet => ({ scenarios: [], current: null, show: null, resultCells: [] })

export function copyScenarios(set: ScenarioSet): ScenarioSet {
  return {
    scenarios: set.scenarios.map((s) => ({ ...s, cells: s.cells.map((c) => ({ ...c })) })),
    current: set.current,
    show: set.show,
    resultCells: set.resultCells.map((r) => ({ ...r })),
  }
}

/** Rows or columns inserted (count > 0) or deleted (count < 0): the cells move; deleted ones leave their scenarios. */
export function shiftScenarios(sheet: Sheet, axis: "rows" | "cols", at: number, count: number) {
  const set = sheet.scenarios
  if (!set.scenarios.length && !set.resultCells.length) return
  const move = (c: ScenarioCell): ScenarioCell | null => {
    const r = shiftRange({ r0: c.row, c0: c.col, r1: c.row, c1: c.col }, axis, at, count)
    return r ? { ...c, row: r.r0, col: r.c0 } : null
  }
  sheet.scenarios = {
    ...set,
    scenarios: set.scenarios.map((s) => ({
      ...s,
      cells: s.cells.map(move).filter((c): c is ScenarioCell => c !== null),
    })),
    resultCells: set.resultCells.map((r) => shiftRange(r, axis, at, count)).filter((r): r is Range => r !== null),
  }
}

/** Show: the scenario's values go into its changing cells (formulas there are replaced, as in Excel). */
export function showScenarioOp(sheet: Sheet, sheetIndex: number, scenario: Scenario): Op {
  const cells: CellWrite[] = scenario.cells.map((c) => {
    const data = { ...(sheet.cells.get(c.row, c.col) ?? {}), value: c.value }
    delete data.input
    if (typeof c.value === "string" && c.value !== "" && /^[=+\-@'\d.]/.test(c.value)) data.input = "'" + c.value
    return { row: c.row, col: c.col, data }
  })
  return { kind: "setCells", sheet: sheetIndex, cells }
}

/** Excel's limit on changing cells per scenario. */
export const MAX_SCENARIO_CELLS = 32
