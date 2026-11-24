import { describe, expect, it, vi } from "vitest"
import { STYLE_NAME_TAKEN, Workbook, type CustomTableStyle, type Selection } from "@cells/core"
import { SheetController } from "./controller"

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
