import { describe, expect, it, vi } from "vitest"
import { pivotChoices, withPageItems, withSort, Workbook, type Selection } from "@cells/core"
import { SheetController } from "./controller"

/** Excel's own run of Show Report Filter Pages had this data (Region as the report filter, "Hidden" unticked). */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data", "Report"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  const rows: (string | number | null)[][] = [
    ["Region", "Product", "Amount"],
    ["East", "Pens", 10],
    ["West", "Pens", 20],
    ["North", "Ink", 5],
    ["East", "Ink", 7],
    ["A/B:C?", "Pens", 3],
    ["Very long region name that goes past thirty one", "Ink", 4],
    [null, "Pens", 2],
    ["West", "Ink", 1],
    ["Data", "Pens", 6],
    ["Hidden", "Ink", 8],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => value !== null && c.sheet.cells.set(row, col, { value })))
  c.pivots.create("Data!A1:C11", { kind: "existing", sheet: 1, at: { row: 4, col: 1 } }, undefined, {
    title: "",
    rows: [1],
    cols: [],
    values: [2],
    pages: [0],
  })
  c.pivots.change((t, cache) => {
    const sorted = withSort(t, cache, 0, "ascending")
    const all = pivotChoices(sorted, cache, 0).items
    return withPageItems(sorted, cache, 0, new Set(all.filter((i) => i.label !== "Hidden").map((i) => i.x)))
  })
  c.wb.clearHistory()
  return c
}

describe("PivotTable Analyze › Options › Show Report Filter Pages", () => {
  it("makes a sheet for each ticked item before the PivotTable's, named as Excel names them", () => {
    const c = setup()
    expect(c.pivots.reportFilters()).toEqual([{ field: 0, name: "Region" }])
    expect(c.pivots.showReportFilterPages(0)).toBe(true)
    // Excel 365 gave exactly these sheets, PivotTable names and order.
    expect(c.wb.sheets.map((s) => s.name)).toEqual([
      "Data",
      "Sheet9",
      "Data (2)",
      "East",
      "North",
      "Sheet5",
      "West",
      "(blank)",
      "Report",
    ])
    expect(c.wb.sheets.slice(1, 8).map((s) => s.pivots[0].name)).toEqual([
      "PivotTable8",
      "PivotTable7",
      "PivotTable6",
      "PivotTable5",
      "PivotTable4",
      "PivotTable3",
      "PivotTable2",
    ])
    expect(c.sheetIndex).toBe(1)
    const east = c.wb.sheets[3]
    const t = east.pivots[0]
    // At A1: the report filter on row 1, the table from row 3, showing East only.
    expect(t.at).toEqual({ row: 2, col: 0 })
    expect(east.cells.get(0, 0)?.value).toBe("Region")
    expect(east.cells.get(0, 1)?.value).toBe("East")
    expect(t.fields[0].multipleItems).toBe(false)
    expect(east.cells.get(5, 1)?.value).toBe(17)
    // The copies share the cache; the PivotTable itself is unchanged.
    expect(c.wb.pivotCaches).toHaveLength(1)
    expect(c.wb.sheets[8].pivots[0].fields[0].multipleItems).toBe(true)
    c.undo()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Data", "Report"])
  })
})
