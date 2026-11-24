import { describe, expect, it, vi } from "vitest"
import { CALC_ITEM_FUNCTIONS, CALC_ITEM_NOT_FIELD, CALC_ITEM_SAME_FIELD, Workbook, type Selection } from "@cells/core"
import { SheetController } from "./controller"

/** Product, Region, Amount on Data!A1:C7; Sum of Amount by Product (rows) and Region (columns) at E1. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
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
    cols: [1],
    values: [2],
  })
  c.wb.clearHistory()
  return c
}
const at = (c: SheetController, row: number, col: number) => c.goTo({ r0: row, c0: col, r1: row, c1: col })
const line = (c: SheetController, label: string) => {
  for (let r = 0; r < 20; r++)
    if (c.sheet.cells.get(r, 4)?.value === label) return [5, 6, 7, 8].map((col) => c.sheet.cells.get(r, col)?.value)
  return null
}

describe("Calculated Item, Solve Order and List Formulas", () => {
  it("adds an item to the field of the active label, worked out in every cell, as one undo step", () => {
    const c = setup()
    at(c, 5, 5)
    expect(c.pivots.calcItemField()).toBe(CALC_ITEM_NOT_FIELD)
    at(c, 3, 4)
    expect(c.pivots.calcItemField()).toBe(0)
    expect(c.pivots.setCalcItem(0, "Writing", "=Pens+Ink*2")).toBeNull()
    // East 10 + 2×2, West 5 + 40×2, total 99 (Excel's).
    expect(line(c, "Writing")).toEqual([14, 85, 99, undefined])
    expect(line(c, "Grand Total")).toEqual([52, 130, 182, undefined])
    expect(c.pivots.calcItemList()).toEqual([
      { field: 0, x: 4, name: "Writing", fieldName: "Product", formula: "=Pens+Ink*2" },
    ])
    c.undo()
    expect(line(c, "Writing")).toBeNull()
    expect(c.wb.pivotCaches[0].calcItems).toBeUndefined()
  })

  it("works out crossings by the solve order, which Move Up changes; Delete takes the item away", () => {
    const c = setup()
    c.pivots.setCalcItem(0, "Both", "=Pens*Ink")
    c.pivots.setCalcItem(1, "Diff", "=East-West")
    expect(line(c, "Both")?.[2]).toBe(-180)
    expect(c.pivots.moveCalcItem(1, -1)).toBe(true)
    expect(c.pivots.calcItemList().map((x) => x.name)).toEqual(["Diff", "Both"])
    expect(line(c, "Both")?.[2]).toBe(-190)
    c.pivots.deleteCalcItem(0, 4)
    expect(line(c, "Both")).toBeNull()
    expect(c.wb.pivotCaches[0].fields[0].items).toEqual(["Pens", "Ink", "Paper", "Clips"])
  })

  it("keeps Excel's rules: no Average and no field summarized twice", () => {
    const c = setup()
    at(c, 2, 5)
    c.pivots.setValueSettings(0, { name: "Average of Amount", fn: "average" })
    expect(c.pivots.setCalcItem(0, "Writing", "=Pens")).toBe(CALC_ITEM_FUNCTIONS)
    c.undo()
    c.pivots.setCalcItem(0, "Writing", "=Pens")
    at(c, 2, 5)
    c.pivots.setValueSettings(0, { name: "Average of Amount", fn: "average" })
    expect(c.shell.alert).toHaveBeenCalledWith(CALC_ITEM_FUNCTIONS)
    expect(c.pivots.active()?.values[0].fn).toBe("sum")
    c.pivots.change((t) => ({ ...t, values: [...t.values, { ...t.values[0], name: "Again" }] }))
    expect(c.shell.alert).toHaveBeenCalledWith(CALC_ITEM_SAME_FIELD)
  })

  it("List Formulas makes a sheet as Excel's", () => {
    const c = setup()
    c.pivots.setCalcField("Tax", "=Amount*0.2")
    c.pivots.setCalcItem(1, "Diff", "=East-West")
    c.pivots.setCalcItem(0, "Writing", "=Pens+Ink*2")
    expect(c.pivots.listFormulas()).toBe(true)
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet2", "Data"])
    expect(c.sheetIndex).toBe(0)
    const text = Array.from({ length: 15 }, (_, r) =>
      [0, 1, 2].map((col) => c.sheet.cells.get(r, col)?.value ?? "").join("|"),
    )
    expect(text.slice(0, 9)).toEqual([
      "Calculated Field||",
      "Solve Order|Field|Formula",
      "1|Tax|=Amount*0.2",
      "||",
      "Calculated Item||",
      "Solve Order|Item|Formula",
      "1|Diff|=East-West",
      "2|Writing|=Pens+Ink*2",
      "||",
    ])
    expect(text[10]).toBe("Note:|When a cell is updated by more than one formula,|")
    expect(c.wb.styles.get(c.sheet.cells.get(1, 2)?.style ?? 0).bold).toBe(true)
    c.undo()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Data"])
  })
})
