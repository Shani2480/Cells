import { describe, expect, it, vi } from "vitest"
import {
  pivotReportLayout,
  withFieldIn,
  withReportLayout,
  withSubtotalsShown,
  Workbook,
  type Selection,
} from "@cells/core"
import { SheetController } from "./controller"

/** Units by Region and Channel (rows) at F1, from Data!A1:D5; G1 and beyond on Data hold nothing else. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data", "Other"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Units", "Amount"],
    ["North", "Shop", 4, 40],
    ["South", "Online", 2, 20],
    ["North", "Online", 5, 50],
    ["East", "Shop", 1, 10],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.pivots.create("A1:D5", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
  c.pivots.change((t, k) => withFieldIn(withFieldIn(withFieldIn(t, k, 0, "rows"), k, 1, "rows"), k, 2, "values"))
  return c
}

const col = (c: SheetController, col: number, rows = 12) =>
  Array.from({ length: rows }, (_, r) => c.sheet.cells.get(r, col)?.value)

describe("PivotTable Analyze", () => {
  it("renames the PivotTable and the active field", () => {
    const c = setup()
    expect(c.pivots.rename("Sales")).toBeNull()
    expect(c.sheet.pivots[0].name).toBe("Sales")
    expect(c.pivots.rename("")).toBe("PivotTable name is not valid.")
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    expect(c.pivots.activeField()?.name).toBe("Region")
    expect(c.pivots.renameActiveField("Area")).toBeNull()
    expect(c.sheet.pivots[0].fields[0].name).toBe("Area")
    c.goTo({ r0: 1, c0: 6, r1: 1, c1: 6 })
    expect(c.pivots.activeField()?.name).toBe("Sum of Units")
    expect(c.pivots.renameActiveField("Region")).toBe("PivotTable field name already exists.")
  })

  it("Clear Filters shows every item again; Clear All leaves it empty", () => {
    const c = setup()
    c.pivots.showItems(0, 5, 0, new Set([0]))
    expect(col(c, 5)).not.toContain("South")
    c.pivots.clear("filters")
    expect(col(c, 5)).toContain("South")
    c.pivots.clear("all")
    expect(c.sheet.pivots[0].rows).toEqual([])
    expect(col(c, 5).filter((v) => v !== undefined)).toEqual([])
    c.undo()
    expect(col(c, 5)).toContain("South")
  })

  it("selects its values, labels and the whole table", () => {
    const c = setup()
    // F1 Row Labels | Sum of Units; F2.. the lines; Grand Total at row 9.
    c.pivots.select("values")
    expect(c.selection.ranges).toEqual([{ r0: 1, c0: 6, r1: 8, c1: 6 }])
    c.pivots.select("labels")
    expect(c.selection.ranges[0]).toEqual({ r0: 0, c0: 5, r1: 8, c1: 5 })
    c.pivots.select("entire")
    expect(c.selection.ranges).toEqual([{ r0: 0, c0: 5, r1: 8, c1: 6 }])
  })

  it("moves the PivotTable on its sheet and to another, in one undo step", () => {
    const c = setup()
    expect(c.pivots.move({ kind: "existing", sheet: 0, at: { row: 20, col: 1 } })).toBeNull()
    expect(c.sheet.pivots[0].at).toEqual({ row: 20, col: 1 })
    expect(c.sheet.cells.get(20, 1)?.value).toBe("Row Labels")
    expect(c.sheet.cells.get(0, 5)).toBeUndefined()
    c.pivots.move({ kind: "existing", sheet: 1, at: { row: 2, col: 0 } })
    expect(c.sheet.name).toBe("Other")
    expect(c.sheet.cells.get(2, 0)?.value).toBe("Row Labels")
    expect(c.wb.sheets[0].pivots).toEqual([])
    expect(c.wb.sheets[0].cells.get(20, 1)).toBeUndefined()
    c.undo()
    expect(c.wb.sheets[0].pivots).toHaveLength(1)
    expect(c.wb.sheets[1].pivots).toHaveLength(0)
    expect(c.wb.sheets[0].cells.get(20, 1)?.value).toBe("Row Labels")
  })

  it("Change Data Source reads the new range, keeping the fields it still has", () => {
    const c = setup()
    expect(c.pivots.sourceText()).toBe("Data!$A$1:$D$5")
    expect(c.pivots.changeSource("Data!A1:D3")).toBeNull()
    expect(c.wb.pivotCaches[0].source).toEqual({ kind: "range", sheet: "Data", ref: "A1:D3" })
    expect(col(c, 5).filter((v) => v !== undefined)).toEqual([
      "Row Labels",
      "North",
      "Shop",
      "South",
      "Online",
      "Grand Total",
    ])
    expect(c.pivots.changeSource("Nowhere!A1")).toBe("Reference isn't valid.")
  })
})

describe("Design", () => {
  it("Report Layout, Subtotals and Grand Totals change every field", () => {
    const c = setup()
    c.pivots.layout((t) => withReportLayout(t, "tabular"), "Report Layout")
    expect(pivotReportLayout(c.sheet.pivots[0])).toBe("tabular")
    // Tabular: Region and Channel in their own columns.
    expect([5, 6, 7].map((k) => c.sheet.cells.get(0, k)?.value)).toEqual(["Region", "Channel", "Sum of Units"])
    c.pivots.layout((t) => withSubtotalsShown(t, "none"), "Subtotals")
    expect(col(c, 5)).not.toContain("North Total")
    c.pivots.setStyleOption("rowStripes", true)
    expect(c.sheet.pivots[0].style.rowStripes).toBe(true)
    c.pivots.setStyle("PivotStyleMedium9")
    expect(c.sheet.pivots[0].style.name).toBe("PivotStyleMedium9")
  })
})
