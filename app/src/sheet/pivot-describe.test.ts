import { describe, expect, it, vi } from "vitest"
import { Workbook, type Selection } from "@cells/core"
import { SheetController } from "./controller"

/** Region, Date, Amount on Data!A1:C6 and a PivotTable of it (Sum of Amount by Region) on Sheet2. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  // 2025-01-10, 2025-04-12, 2025-07-15, 2025-01-20, 2025-10-02 as serials.
  const rows: (string | number)[][] = [
    ["Region", "Date", "Amount"],
    ["North", 45667, 40],
    ["South", 45759, 25],
    ["North", 45853, 50],
    ["East", 45677, 10],
    ["South", 45932, 70],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.goTo({ r0: 1, c0: 1, r1: 5, c1: 1 })
  c.format.numberFormat("yyyy-mm-dd")
  c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
  c.pivots.create("Data!$A$1:$C$6", { kind: "newSheet" }, undefined, { title: "", rows: [0], cols: [], values: [2] })
  return c
}

const column = (c: SheetController, col: number, from: number, to: number) =>
  Array.from({ length: to - from + 1 }, (_, i) => c.sheet.cells.get(from + i, col)?.value)

describe("Describe the PivotTable you want (Fields pane)", () => {
  it("lays the active PivotTable out anew as described, in one undo step, keeping its name and place", () => {
    const c = setup()
    const before = c.sheet.pivots[0]
    expect(
      c.pivots.applyRecipe({
        title: "",
        rows: [0],
        cols: [],
        values: [2],
        sort: { field: 0, by: 0, order: "descending" },
      }),
    ).toBe(true)
    const after = c.sheet.pivots[0]
    expect(after).toMatchObject({ name: before.name, at: before.at })
    expect(after.fields[0]).toMatchObject({ sort: "descending", sortBy: 0 })
    // South 95, North 90, East 10.
    expect(column(c, 0, 3, 6)).toEqual(["South", "North", "East", "Grand Total"])
    c.undo()
    expect(c.sheet.pivots[0]).toEqual(before)
  })

  it("groups dates by quarter in the PivotTable's cache, the quarters taking the date's place", () => {
    const c = setup()
    const label = c.pivots.applyRecipe({
      title: "",
      rows: [1],
      cols: [],
      values: [2],
      dates: { field: 1, parts: ["quarters"] },
    })
    expect(label).toBe(true)
    const [table] = c.sheet.pivots
    const cache = c.wb.pivotCaches.find((x) => x.id === table.cacheId)!
    expect(table.rows.map((f) => cache.fields[f].name)).toEqual(["Quarters (Date)"])
    expect(column(c, 0, 3, 7)).toEqual(["Qtr1", "Qtr2", "Qtr3", "Qtr4", "Grand Total"])
    expect(column(c, 1, 3, 7)).toEqual([50, 25, 50, 70, 195])
    c.undo()
    expect(c.wb.pivotCaches[0].fields[1].group).toBeUndefined()
  })
})
