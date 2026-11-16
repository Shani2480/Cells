import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { nameFromLabel } from "./names"

function setup(cells: Record<string, string>) {
  const c = new SheetController()
  for (const [ref, text] of Object.entries(cells)) {
    const col = ref.charCodeAt(0) - 65
    const row = Number(ref.slice(1)) - 1
    c.goTo({ r0: row, c0: col, r1: row, c1: col })
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  c.wb.calculate(true)
  return c
}
const at = (c: SheetController, ref: string) => {
  const col = ref.charCodeAt(0) - 65
  const row = Number(ref.slice(1)) - 1
  c.goTo({ r0: row, c0: col, r1: row, c1: col })
}
const arrows = (c: SheetController) =>
  c.audit
    .arrows()
    .map(
      (a) =>
        `${a.from ? `${a.from.r0},${a.from.c0}:${a.from.r1},${a.from.c1}` : "sheet"}>${a.to.row},${a.to.col}${a.error ? "!" : ""}`,
    )
    .sort()

describe("Trace Precedents and Dependents", () => {
  it("go one level further with each click and follow edits; Remove Arrows clears them", () => {
    const c = setup({ A1: "1", A2: "2", B1: "=A1+A2", C1: "=B1*2", D1: "=SUM(A1:A2)+Sheet2!A1" })
    at(c, "C1")
    expect(c.audit.tracePrecedents()).toBe(true)
    expect(arrows(c)).toEqual(["0,1:0,1>0,2"])
    expect(c.audit.tracePrecedents()).toBe(true)
    expect(arrows(c)).toEqual(["0,0:0,0>0,1", "0,1:0,1>0,2", "1,0:1,0>0,1"])
    expect(c.audit.tracePrecedents()).toBe(false)
    at(c, "D1")
    c.audit.tracePrecedents()
    expect(arrows(c)).toContain("0,0:1,0>0,3")
    expect(arrows(c)).toContain("sheet>0,3")
    c.audit.removeArrows("precedents")
    at(c, "A1")
    expect(c.audit.traceDependents()).toBe(true)
    expect(arrows(c)).toEqual(["0,0:0,0>0,1", "0,0:0,0>0,3"])
    expect(c.audit.traceDependents()).toBe(true)
    expect(arrows(c)).toContain("0,1:0,1>0,2")
    c.audit.removeArrows()
    expect(c.audit.hasArrows).toBe(false)
  })

  it("Trace Error follows the error back to where it starts, in red", () => {
    const c = setup({ A1: "0", B1: "=1/A1", C1: "=B1+1", D1: "5" })
    at(c, "D1")
    expect(c.audit.traceError()).toBe(false)
    at(c, "C1")
    expect(c.audit.traceError()).toBe(true)
    expect(arrows(c)).toEqual(["0,0:0,0>0,1", "0,1:0,1>0,2!"])
  })
})

describe("Error Checking", () => {
  it("steps through errors and numbers stored as text; Ignore Error and Convert to Number", () => {
    const c = setup({ A1: "0", B2: "=1/A1", C3: "'42", D4: "text" })
    expect(c.audit.problemAt(1, 1)).toEqual({ kind: "error", code: "#DIV/0!" })
    expect(c.audit.problemAt(2, 2)).toEqual({ kind: "numberAsText" })
    expect(c.audit.problemAt(3, 3)).toBeNull()
    expect(c.audit.nextProblem({ row: 0, col: 0 })).toEqual({ row: 1, col: 1 })
    expect(c.audit.nextProblem({ row: 1, col: 1 })).toEqual({ row: 2, col: 2 })
    expect(c.audit.nextProblem({ row: 2, col: 2 })).toEqual({ row: 1, col: 1 })
    expect(c.audit.nextProblem({ row: 1, col: 1 }, true)).toEqual({ row: 2, col: 2 })
    c.audit.convertToNumber({ row: 2, col: 2 })
    expect(c.sheet.cells.value(2, 2)).toBe(42)
    c.audit.ignoreError({ row: 1, col: 1 })
    expect(c.audit.nextProblem({ row: 0, col: 0 })).toBeNull()
    c.audit.resetIgnored()
    expect(c.audit.nextProblem({ row: 0, col: 0 })).toEqual({ row: 1, col: 1 })
  })
})

describe("Watch Window", () => {
  it("adds the selected cells once each and deletes them", () => {
    const c = setup({ A1: "1", A2: "2" })
    c.goTo({ r0: 0, c0: 0, r1: 1, c1: 0 })
    c.audit.addWatches()
    c.audit.addWatches()
    expect(c.audit.watches.map((w) => [w.row, w.col])).toEqual([
      [0, 0],
      [1, 0],
    ])
    c.audit.deleteWatch(0)
    expect(c.audit.watches.map((w) => w.row)).toEqual([1])
  })
})

describe("Create from Selection and Delete", () => {
  it("names each column from its header and each row from its label", () => {
    const c = setup({ B1: "Jan", C1: "Feb sales", A2: "Rent", A3: "2024", B2: "10", C2: "20", B3: "1", C3: "2" })
    c.goTo({ r0: 0, c0: 0, r1: 2, c1: 2 })
    expect(c.names.guessEdges()).toEqual({ top: true, left: false, bottom: false, right: false })
    expect(c.names.createFromSelection({ top: true, left: true, bottom: false, right: false })).toBeNull()
    expect(c.names.all().map((n) => [n.name, n.ref])).toEqual(
      expect.arrayContaining([
        ["Jan", "Sheet1!$B$2:$B$3"],
        ["Feb_sales", "Sheet1!$C$2:$C$3"],
        ["Rent", "Sheet1!$B$2:$C$2"],
        ["_2024", "Sheet1!$B$3:$C$3"],
      ]),
    )
    expect(c.names.all()).toHaveLength(4)
    c.names.remove(["Jan", "rent"])
    expect(
      c.names
        .all()
        .map((n) => n.name)
        .sort(),
    ).toEqual(["Feb_sales", "_2024"])
    expect([nameFromLabel("A1"), nameFromLabel("q-1 total"), nameFromLabel("")]).toEqual(["_A1", "q_1_total", null])
  })
})
