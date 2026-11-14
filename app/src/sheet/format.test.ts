import { parseRange, selectRange, writeInputOp } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { NUMBER_FORMATS } from "./format"

// Expectations follow Excel's own ribbon commands, run in Excel through CommandBars.ExecuteMso.

function setup() {
  const c = new SheetController()
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  const style = (ref: string) => {
    const r = parseRange(ref)!
    return c.wb.styles.get(c.sheet.styleAt(r.r0, r.c0))
  }
  const put = (ref: string, text: string) => {
    const r = parseRange(ref)!
    c.wb.transact([writeInputOp(0, c.sheet, c.wb.styles, r.r0, r.c0, text)], { label: "Typing" })
  }
  const key = (key: string, shift = false) => c.keyReady({ key, primary: true, shift, alt: false })
  return { c, f: c.format, select, style, put, key }
}

describe("formatting commands", () => {
  it("Bold follows the active cell: off there turns all on, on there turns all off", () => {
    const { f, select, style } = setup()
    select("A2")
    f.toggle("bold")
    select("A1:A2") // active A1 is not bold
    f.toggle("bold")
    expect(style("A1").bold).toBe(true)
    select("A2:A3") // active A2 is bold
    f.toggle("bold")
    expect(style("A2").bold).toBeUndefined()
    expect(style("A3").bold).toBeUndefined()
  })

  it("Grow Font takes the active cell's size one step and gives it to every cell", () => {
    const { f, select, style } = setup()
    select("A2")
    f.fontSize(20)
    select("A1:A2")
    f.stepFont(1)
    expect(style("A1").fontSize).toBe(12)
    expect(style("A2").fontSize).toBe(12)
    f.stepFont(-1)
    expect(style("A2").fontSize).toBeUndefined() // 11 is the default
  })

  it("Increase Indent turns the alignment to Left, even from Center", () => {
    const { f, select, style } = setup()
    select("B2")
    f.hAlign("center")
    f.indent(1)
    expect(style("B2")).toMatchObject({ hAlign: "left", indent: 1 })
    f.indent(-1)
    expect(style("B2").indent).toBeUndefined()
  })

  it("alignment buttons toggle back to General", () => {
    const { f, select, style } = setup()
    select("C3")
    f.hAlign("right")
    expect(style("C3").hAlign).toBe("right")
    f.hAlign("right")
    expect(style("C3").hAlign).toBeUndefined()
    f.vAlign("top")
    f.vAlign("bottom")
    expect(style("C3").vAlign).toBeUndefined()
  })

  it("Merge & Center warns before discarding values, keeps the top-left one, and undoes in one step", () => {
    const { c, f, select, style, put } = setup()
    put("C3", "a")
    put("D3", "b")
    let asked = ""
    c.shell.confirm = (message, ok) => {
      asked = message
      ok()
    }
    select("C3:D4")
    f.merge("center")
    expect(asked).toMatch(/upper-left value/)
    expect(c.sheet.merges.all).toEqual([parseRange("C3:D4")])
    expect(style("C3").hAlign).toBe("center")
    expect(c.sheet.cells.value(2, 3)).toBeUndefined()
    expect(c.sheet.cells.value(2, 2)).toBe("a")
    expect(f.isMerged()).toBe(true)
    c.undo()
    expect(c.sheet.merges.size).toBe(0)
    expect(c.sheet.cells.value(2, 3)).toBe("b")
  })

  it("Merge Across makes one merge per row; cancelling the warning changes nothing", () => {
    const { c, f, select, put } = setup()
    select("A1:C3")
    f.merge("across")
    expect(c.sheet.merges.size).toBe(3)
    c.undo()
    put("A1", "x")
    put("B1", "y")
    c.shell.confirm = () => undefined
    select("A1:C3")
    f.merge("cells")
    expect(c.sheet.merges.size).toBe(0)
  })

  it("Accounting, Percent and Comma set the format and its cell style", () => {
    const { f, select, style } = setup()
    select("B2")
    f.numberFormat(NUMBER_FORMATS.accounting, "Currency")
    expect(style("B2")).toEqual({ numberFormat: NUMBER_FORMATS.accounting, cellStyle: "Currency" })
    f.numberFormat("General")
    expect(style("B2").numberFormat).toBeUndefined()
  })

  it("Increase Decimal on General starts from what the cell shows", () => {
    const { c, f, select, style, put } = setup()
    put("A1", "3.14159")
    select("A1")
    f.decimals(1)
    expect(style("A1").numberFormat).toBe("0.000000")
    c.shownText = () => "1234567.9"
    put("A2", "1234567.891")
    select("A2")
    f.decimals(1)
    expect(style("A2").numberFormat).toBe("0.00")
  })

  it("Format Painter copies the pattern, tiled over a bigger target", () => {
    const { f, select, style } = setup()
    select("A1")
    f.toggle("bold")
    select("A2")
    f.fill("#ffff00")
    select("A1:A2")
    f.armPainter()
    f.paint(parseRange("C1:C4")!)
    expect(style("C1")).toEqual({ bold: true })
    expect(style("C2")).toEqual({ fill: "#ffff00" })
    expect(style("C3")).toEqual({ bold: true })
    expect(style("C4")).toEqual({ fill: "#ffff00" })
    expect(f.painter).toBeNull()
  })

  it("Clear Formats removes formats and merges, keeps values", () => {
    const { c, f, select, style, put } = setup()
    put("A1", "5")
    select("A1:B1")
    f.toggle("italic")
    f.merge("cells")
    f.clearFormats()
    expect(style("A1")).toEqual({})
    expect(c.sheet.merges.size).toBe(0)
    expect(c.sheet.cells.value(0, 0)).toBe(5)
  })

  it("shortcuts: Ctrl+B, Ctrl+5, Ctrl+Shift+%, Ctrl+Shift+&, Ctrl+Shift+_, Ctrl+1", () => {
    const { c, select, style, key } = setup()
    let dialog = ""
    c.shell.formatCells = (tab) => (dialog = tab ?? "number")
    select("B2")
    key("b")
    key("5")
    key("%", true)
    expect(style("B2")).toMatchObject({ bold: true, strike: true, numberFormat: "0%" })
    key("&", true)
    expect(style("B2").borderTop).toEqual({ style: "thin" })
    key("_", true)
    expect(style("B2").borderTop).toBeUndefined()
    key("1")
    expect(dialog).toBe("number")
    // Ctrl+- still deletes, Ctrl+` still shows formulas.
    key("`")
    expect(c.showFormulas).toBe(true)
  })
})
