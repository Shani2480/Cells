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
