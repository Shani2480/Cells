import { describe, expect, it, vi } from "vitest"
import { Workbook, type Selection } from "@cells/core"
import { SheetController } from "./controller"
import { PIVOT_SORT_MESSAGE } from "./data"

/** Product, Region, Amount on Data!A1:C7; a PivotTable of Sum of Amount by Product (rows), at Data!E1. */
function setup(cols: number[] = []) {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.openDialog = vi.fn()
  const rows: (string | number)[][] = [
    ["Product", "Region", "Amount"],
    ["Pens", "East", 10],
    ["Ink", "West", 40],
    ["Paper", "East", 25],
    ["Pens", "West", 5],
    ["Clips", "East", 1],
    ["Ink", "East", 2],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  c.pivots.create("A1:C7", { kind: "existing", sheet: 0, at: { row: 0, col: 4 } }, undefined, {
    title: "",
    rows: [0],
    cols,
    values: [2],
  })
  c.wb.clearHistory()
  return c
}
const at = (c: SheetController, row: number, col: number) => c.goTo({ r0: row, c0: col, r1: row, c1: col })
const labels = (c: SheetController, from = 1, count = 4) =>
  Array.from({ length: count }, (_, i) => c.sheet.cells.get(from + i, 4)?.value)

describe("Sort in a PivotTable", () => {
  it("sorts a field by its labels from a label, and by a value from a value cell", () => {
    const c = setup()
    expect(labels(c)).toEqual(["Pens", "Ink", "Paper", "Clips"])
    // A value cell: the field of its row, by that value (Ink 42, Paper 25, Pens 15, Clips 1).
    at(c, 2, 5)
    expect(c.pivots.sortTarget()).toEqual({ field: 0, by: 0 })
    c.data.quickSort(true)
    expect(labels(c)).toEqual(["Ink", "Paper", "Pens", "Clips"])
    expect(c.pivots.active()?.fields[0]).toMatchObject({ sort: "descending", sortBy: 0 })
    // A label: by the labels again, which ends the sort by value.
    at(c, 1, 4)
    expect(c.pivots.sortTarget()).toEqual({ field: 0, by: null })
    expect(c.pivots.sortActive(true)).toBe(true)
    expect(labels(c)).toEqual(["Pens", "Paper", "Ink", "Clips"])
    expect(c.pivots.active()?.fields[0].sortBy).toBeUndefined()
    c.undo()
    expect(labels(c)).toEqual(["Ink", "Paper", "Pens", "Clips"])
  })

  it("Data › Sort opens the field's Sort dialog, which sets Manual, labels or a value", () => {
    const c = setup()
    at(c, 2, 5)
    c.data.openSortDialog()
    expect(c.shell.openDialog).toHaveBeenCalledWith("pivotSort:0")
    expect(c.pivots.sortDialog(0)).toEqual({ name: "Product", sort: "manual", by: null, values: ["Sum of Amount"] })
    c.pivots.setSort(0, "ascending", 0)
    expect(labels(c)).toEqual(["Clips", "Pens", "Paper", "Ink"])
    expect(c.pivots.sortDialog(0)).toMatchObject({ sort: "ascending", by: 0 })
    // Manual keeps the items where they are now, free to be moved.
    c.pivots.setSort(0, "manual", null)
    expect(labels(c)).toEqual(["Clips", "Pens", "Paper", "Ink"])
    expect(c.pivots.sortDialog(0)).toMatchObject({ sort: "manual", by: null })
  })

  it("with a column field, sorts by the grand totals only", () => {
    const c = setup([1])
    // E1 Sum of Amount, E2 Row Labels | East | West | Grand Total; products from row 3.
    expect(c.sheet.cells.get(1, 7)?.value).toBe("Grand Total")
    at(c, 3, 5)
    expect(c.pivots.sortTarget()).toBeNull()
    c.data.quickSort(false)
    expect(c.shell.alert).toHaveBeenCalledWith(PIVOT_SORT_MESSAGE)
    // The Grand Total column sorts the products.
    at(c, 3, 7)
    expect(c.pivots.sortTarget()).toEqual({ field: 0, by: 0 })
    c.pivots.sortActive(true)
    expect(labels(c, 2)).toEqual(["Ink", "Paper", "Pens", "Clips"])
    // The Grand Total row sorts the regions (West 45, East 38).
    at(c, 6, 5)
    expect(c.sheet.cells.get(6, 4)?.value).toBe("Grand Total")
    expect(c.pivots.sortTarget()).toEqual({ field: 1, by: 0 })
    c.pivots.sortActive(true)
    expect([5, 6].map((col) => c.sheet.cells.get(1, col)?.value)).toEqual(["West", "East"])
  })
})
