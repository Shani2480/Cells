import { describe, expect, it, vi } from "vitest"
import { STYLE_NAME_TAKEN, Workbook, type CustomTableStyle, type Selection } from "@cells/core"
import { SheetController } from "./controller"
import { OVERLAP_PIVOT } from "./tables"

/** Product and Amount on Data!A1:B4, a PivotTable of them at D3. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  const rows: (string | number)[][] = [
    ["Product", "Amount"],
    ["Pens", 10],
    ["Ink", 40],
    ["Paper", 25],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  c.pivots.create("A1:B4", { kind: "existing", sheet: 0, at: { row: 2, col: 3 } }, undefined, {
    title: "",
    rows: [0],
    cols: [],
    values: [1],
  })
  c.wb.clearHistory()
  c.goTo({ r0: 3, c0: 3, r1: 3, c1: 3 })
  return c
}

const mine = (name = "PivotTable Style 1"): CustomTableStyle => ({
  name,
  table: false,
  pivot: true,
  elements: { headerRow: { bold: true, fill: "#1F3F7F", color: "#FFFFFF" } },
})
const styleOf = (c: SheetController) => c.pivots.active()!.style.name

describe("New PivotTable Style", () => {
  it("saves a style as one undo step, refuses a taken name, and new PivotTables take the default", () => {
    const c = setup()
    expect(c.pivotStyles.save({ style: mine("PivotStyleLight16"), isDefault: false })).toBe(STYLE_NAME_TAKEN)
    expect(c.pivotStyles.save({ style: mine(), isDefault: true })).toBeNull()
    expect(c.pivotStyles.list().map((s) => s.name)).toEqual(["PivotTable Style 1"])
    expect(c.pivotStyles.save({ style: mine("pivottable style 1"), isDefault: false })).toBe(STYLE_NAME_TAKEN)
    expect(c.pivotStyles.defaultStyle()).toBe("PivotTable Style 1")
    c.wb.undo()
    expect(c.pivotStyles.list()).toEqual([])
    expect(c.pivotStyles.defaultStyle()).toBe("PivotStyleLight16")
    c.wb.redo()
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.pivots.create("A1:B4", { kind: "newSheet" }, undefined, { title: "", rows: [0], cols: [], values: [1] })
    expect(c.wb.sheets[c.sheetIndex].pivots[0].style.name).toBe("PivotTable Style 1")
  })

  it("Modify renames the style and the PivotTables with it; Delete puts them back on Light 16", () => {
    const c = setup()
    c.pivotStyles.save({ style: mine(), isDefault: true })
    c.pivots.setStyle("PivotTable Style 1")
    expect(c.pivotStyles.save({ style: mine("Sales"), isDefault: true }, "PivotTable Style 1")).toBeNull()
    expect(styleOf(c)).toBe("Sales")
    expect(c.pivotStyles.defaultStyle()).toBe("Sales")
    c.wb.undo()
    expect(styleOf(c)).toBe("PivotTable Style 1")
    c.wb.redo()
    expect(c.pivotStyles.remove("Sales")).toBe(true)
    expect(styleOf(c)).toBe("PivotStyleLight16")
    expect(c.pivotStyles.defaultStyle()).toBe("PivotStyleLight16")
    c.wb.undo()
    expect(styleOf(c)).toBe("Sales")
    expect(c.pivotStyles.get("sales")?.elements.headerRow?.fill).toBe("#1F3F7F")
  })

  it("Set As Default takes a built-in style too", () => {
    const c = setup()
    expect(c.pivotStyles.setDefault("PivotStyleMedium9")).toBe(true)
    expect(c.wb.tableStyles.defaultPivot).toBe("PivotStyleMedium9")
    expect(c.pivotStyles.setDefault("PivotStyleLight16")).toBe(true)
    expect(c.wb.tableStyles.defaultPivot).toBeNull()
  })
})

describe("New Table Style", () => {
  /** Item and Qty on Data!A1:B4 as a table. */
  function withTable() {
    const c = new SheetController(new Workbook<Selection>(["Data"]))
    c.shell.alert = vi.fn()
    ;[
      ["Item", "Qty"],
      ["Bolts", 3],
      ["Nuts", 6],
      ["Washers", 9],
    ].forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
    c.tables.create({ r0: 0, c0: 0, r1: 3, c1: 1 }, true, "TableStyleMedium2")
    c.wb.clearHistory()
    c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
    return c
  }
  const table = (name = "Parts"): CustomTableStyle => ({
    name,
    table: true,
    pivot: false,
    elements: { headerRow: { fill: "#1F3F7F" }, firstRowStripe: { fill: "#DAE6F2", stripe: 2 } },
  })

  it("offers a table style only to tables, shares names with PivotTable styles, and new tables take the default", () => {
    const c = withTable()
    expect(c.tableStyles.save({ style: table(), isDefault: true })).toBeNull()
    expect(c.tableStyles.list().map((s) => s.name)).toEqual(["Parts"])
    expect(c.pivotStyles.list()).toEqual([])
    expect(c.pivotStyles.save({ style: mine("parts"), isDefault: false })).toBe(STYLE_NAME_TAKEN)
    expect(c.tableStyles.defaultStyle()).toBe("Parts")
    // Insert › Table / Ctrl+T offers the default style.
    c.shell.createTable = vi.fn()
    c.goTo({ r0: 10, c0: 0, r1: 12, c1: 1 })
    c.tables.openCreate()
    expect(c.shell.createTable).toHaveBeenCalledWith(expect.objectContaining({ style: "Parts" }))
  })

  it("Modify renames the tables with the style; Delete puts them back on Medium 2, as Excel", () => {
    const c = withTable()
    c.tableStyles.save({ style: table(), isDefault: true })
    c.tables.setStyle("Parts")
    expect(c.tableStyles.save({ style: table("Stock"), isDefault: true }, "Parts")).toBeNull()
    expect(c.sheet.tables[0].style).toBe("Stock")
    expect(c.tableStyles.remove("Stock")).toBe(true)
    expect(c.sheet.tables[0].style).toBe("TableStyleMedium2")
    expect(c.tableStyles.defaultStyle()).toBe("TableStyleMedium2")
    c.wb.undo()
    expect(c.sheet.tables[0].style).toBe("Stock")
    expect(c.wb.tableStyles.defaultTable).toBe("Stock")
  })
})

describe("a table over a PivotTable", () => {
  it("is refused with Excel's message", () => {
    const c = setup()
    c.shell.alert = vi.fn()
    expect(c.tables.create({ r0: 3, c0: 3, r1: 5, c1: 4 }, true, "TableStyleMedium2")).toBe(false)
    expect(c.shell.alert).toHaveBeenCalledWith(OVERLAP_PIVOT)
    expect(c.sheet.tables).toEqual([])
  })
})
