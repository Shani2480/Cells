import { describe, expect, it } from "vitest"
import { scenarioSummary, scenarioSummaryCells, showScenarioOp, Workbook, type ScenarioSet } from "../src"

function book() {
  const wb = new Workbook()
  wb.transact(
    [
      {
        kind: "setCells",
        sheet: 0,
        cells: [
          { row: 0, col: 0, data: { value: 10 } },
          { row: 1, col: 0, data: { value: 2 } },
          { row: 2, col: 0, data: { input: "=A1*A2", value: 0 } },
        ],
      },
    ],
    { label: "Load", history: false },
  )
  wb.calculate(true)
  return wb
}

const set = (): ScenarioSet => ({
  scenarios: [
    {
      name: "Best",
      comment: "",
      user: "Alex",
      locked: true,
      hidden: false,
      cells: [
        { row: 0, col: 0, value: 20 },
        { row: 1, col: 0, value: 3 },
      ],
    },
    { name: "Worst", comment: "", user: "Alex", locked: true, hidden: false, cells: [{ row: 0, col: 0, value: 5 }] },
  ],
  current: null,
  show: null,
  resultCells: [],
})

describe("Scenarios", () => {
  it("are kept on the sheet with undo", () => {
    const wb = book()
    wb.transact([{ kind: "scenarios", sheet: 0, scenarios: set() }], { label: "Add Scenario" })
    expect(wb.sheets[0].scenarios.scenarios.map((s) => s.name)).toEqual(["Best", "Worst"])
    wb.undo()
    expect(wb.sheets[0].scenarios.scenarios).toEqual([])
  })

  it("move with inserted rows and lose deleted cells, undone with the edit", () => {
    const wb = book()
    wb.transact([{ kind: "scenarios", sheet: 0, scenarios: set() }], { label: "Add Scenario" })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 1, count: 2 }], { label: "Insert" })
    expect(wb.sheets[0].scenarios.scenarios[0].cells.map((c) => c.row)).toEqual([0, 3])
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 0, count: 1 }], { label: "Delete" })
    expect(wb.sheets[0].scenarios.scenarios[0].cells).toEqual([{ row: 2, col: 0, value: 3 }])
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].scenarios.scenarios[0].cells.map((c) => c.row)).toEqual([0, 1])
  })

  it("Show puts the values in the changing cells", () => {
    const wb = book()
    const s = set().scenarios[0]
    wb.transact([showScenarioOp(wb.sheets[0], 0, s)], { label: "Show" })
    expect(wb.sheets[0].cells.value(2, 0)).toBe(60)
  })

  it("Summary works out the result cells for each scenario, leaving the sheet as it is", () => {
    const wb = book()
    wb.transact([{ kind: "scenarios", sheet: 0, scenarios: set() }], { label: "Add Scenario" })
    const summary = scenarioSummary(wb, 0, [{ r0: 2, c0: 0, r1: 2, c1: 0 }])
    expect(summary.names).toEqual(["Best", "Worst"])
    expect(summary.changing).toEqual([
      { ref: "$A$1", current: 10, values: [20, 5] },
      { ref: "$A$2", current: 2, values: [3, undefined] },
    ])
    expect(summary.results).toEqual([{ ref: "$A$3", current: 20, values: [60, 10] }])
    expect(wb.sheets[0].cells.value(2, 0)).toBe(20)
    const laid = scenarioSummaryCells(summary)
    const at = (row: number, col: number) => laid.cells.find((c) => c.row === row && c.col === col)?.data?.value
    expect([at(1, 1), at(1, 3), at(1, 4), at(1, 5)]).toEqual(["Scenario Summary", "Current Values:", "Best", "Worst"])
    expect([at(5, 1), at(6, 2), at(6, 4)]).toEqual(["Result Cells:", "$A$3", 60])
    expect(laid.table).toEqual({ r0: 1, c0: 1, r1: 6, c1: 5 })
  })
})
