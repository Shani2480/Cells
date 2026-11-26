import { parseRange, selectRange } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { NAME_BOX_INVALID, absoluteRef } from "./names"

function setup() {
  const c = new SheetController()
  const messages: string[] = []
  c.shell = { ...c.shell, alert: (m) => messages.push(m) }
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  return { c, select, messages }
}

describe("names in the Name Box", () => {
  it("typing a new name names the selection; the Name Box then shows it", () => {
    const { c, select } = setup()
    select("B2:C4")
    c.names.fromNameBox("Budget")
    expect(c.wb.names.get("Budget")?.ref).toBe("Sheet1!$B$2:$C$4")
    expect(c.nameBoxText()).toBe("Budget")
    select("A1")
    expect(c.nameBoxText()).toBe("A1")
    c.names.fromNameBox("budget")
    expect(c.selection.ranges[0]).toEqual({ r0: 1, c0: 1, r1: 3, c1: 2 })
  })

  it("goes to a reference on another sheet", () => {
    const { c } = setup()
    c.sheets.insert(1)
    c.sheets.rename(1, "Q2 Data")
    c.activate(0)
    c.names.fromNameBox("'Q2 Data'!D5")
    expect(c.sheetIndex).toBe(1)
    expect(c.selection.active).toEqual({ row: 4, col: 3 })
  })

  it("refuses text that is neither a reference nor a name", () => {
    const { c, messages } = setup()
    c.names.fromNameBox("two words")
    expect(messages).toEqual([NAME_BOX_INVALID])
    expect(c.wb.names.all()).toHaveLength(0)
  })

  it("Define Name reports Excel's problems", () => {
    const { c } = setup()
    expect(c.names.define("A1", "=Sheet1!$A$1")).toMatch(/not valid/)
    expect(c.names.define("Rate", "=Sheet1!$A$1")).toBeNull()
    expect(c.names.define("rate", "=Sheet1!$A$2")).toMatch(/already exists/)
  })

  it("goes to a list of ranges, selecting them all, the last one's first cell active (as Excel)", () => {
    const { c } = setup()
    c.names.fromNameBox("B1, D3:E4")
    expect(c.selection.ranges).toEqual([
      { r0: 0, c0: 1, r1: 0, c1: 1 },
      { r0: 2, c0: 3, r1: 3, c1: 4 },
    ])
    expect(c.selection.active).toEqual({ row: 2, col: 3 })
    // Semicolons too (the list separator where the comma is the decimal point).
    c.names.fromNameBox("D3:E4;B1")
    expect(c.selection.active).toEqual({ row: 0, col: 1 })
    expect(c.wb.names.all()).toHaveLength(0)
  })

  it("names a selection of several ranges, goes back to it, and shows its name", () => {
    const { c, select } = setup()
    c.names.fromNameBox("D1,F3:G4")
    c.names.fromNameBox("Multi")
    expect(c.wb.names.get("Multi")?.ref).toBe("Sheet1!$D$1,Sheet1!$F$3:$G$4")
    select("A1")
    c.names.fromNameBox("multi")
    expect(c.selection.ranges).toHaveLength(2)
    expect(c.selection.active).toEqual({ row: 2, col: 5 })
    expect(c.nameBoxText()).toBe("Multi")
  })

  it("goes to a table's data rows, without its header and total rows", () => {
    const { c, select } = setup()
    c.sheet.cells.set(0, 0, { value: "Item" })
    c.sheet.cells.set(0, 1, { value: "Qty" })
    c.sheet.cells.set(1, 0, { value: "Pens" })
    c.sheet.cells.set(2, 0, { value: "Ink" })
    c.tables.create({ r0: 0, c0: 0, r1: 2, c1: 1 }, true, "TableStyleMedium2")
    const name = c.sheet.tables[0].name
    select("A2")
    c.tables.setOption("totals", true)
    select("F9")
    c.names.fromNameBox(name)
    expect(c.selection.ranges[0]).toEqual({ r0: 1, c0: 0, r1: 2, c1: 1 })
  })

  it("refuses a list across sheets", () => {
    const { c, messages } = setup()
    c.sheets.insert(1)
    c.activate(0)
    c.names.fromNameBox("A1,Sheet2!B2")
    expect(messages).toEqual([NAME_BOX_INVALID])
  })

  it("lists the names that point at cells and the tables, A to Z", () => {
    const { c } = setup()
    c.names.define("rate", "=0.2")
    c.names.define("Zones", "=Sheet1!$A$1:$A$5")
    c.names.define("budget", "=Sheet1!$B$1,Sheet1!$C$3")
    c.sheet.cells.set(0, 4, { value: "Head" })
    c.tables.create({ r0: 0, c0: 4, r1: 2, c1: 4 }, true, "TableStyleMedium2")
    const table = c.sheet.tables[0].name
    expect(c.names.boxNames()).toEqual(["budget", table, "Zones"].sort((a, b) => a.localeCompare(b)))
  })

  it("writes whole rows and columns absolutely", () => {
    expect(absoluteRef({ r0: 0, r1: 1_048_575, c0: 0, c1: 2 })).toBe("$A:$C")
    expect(absoluteRef({ r0: 2, r1: 4, c0: 0, c1: 16_383 })).toBe("$3:$5")
  })
})

describe("insert function", () => {
  it("starts =NAME() with the caret inside, or inserts at the caret of an entry", () => {
    const { c } = setup()
    c.insertFunction("SUM")
    expect(c.edit?.text).toBe("=SUM()")
    expect(c.edit?.selStart).toBe(5)
    c.setEditText("=1+", 3, 3)
    c.insertFunction("ABS")
    expect(c.edit?.text).toBe("=1+ABS()")
    expect(c.edit?.selStart).toBe(7)
  })

  it("Ctrl+Shift+U expands the formula bar", () => {
    const { c } = setup()
    c.keyReady({ key: "U", primary: true, shift: true, alt: false })
    expect(c.formulaBarExpanded).toBe(true)
    c.setFormulaBarLines(50)
    expect(c.formulaBarLines).toBe(20)
  })
})
