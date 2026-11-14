import { parseRange, selectRange } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { fittedRowHeight } from "./autofit"

/** 7 px per character at 11 pt, scaled with the font size. */
const measure = (text: string, font?: string) => {
  const px = font ? Number(/([\d.]+)px/.exec(font)?.[1] ?? 14.67) : 14.67
  return text.length * 7 * (px / 14.67)
}

function setup() {
  const c = new SheetController()
  c.rowFit = (row) => fittedRowHeight(c, row, measure)
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  const type = (ref: string, text: string) => {
    select(ref)
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  return { c, select, type }
}

describe("rows fit their content", () => {
  it("Alt+Enter lines turn on wrap and heighten the row; undo takes both back", () => {
    const { c, type } = setup()
    type("A2", "first\nsecond\nthird")
    expect(c.wb.styles.get(c.sheet.styleAt(1, 0)).wrap).toBe(true)
    expect(c.sheet.rows.size(1)).toBe(60)
    expect(c.sheet.rows.isAuto(1)).toBe(true)
    c.undo()
    expect(c.sheet.rows.size(1)).toBe(20)
  })

  it("a bigger font heightens the row, and going back lowers it", () => {
    const { c, select, type } = setup()
    type("B3", "Title")
    select("B3")
    c.format.apply({ fontSize: 24 }, "Font Size")
    expect(c.sheet.rows.size(2)).toBe(Math.round(32 * 1.36))
    c.format.apply({ fontSize: null }, "Font Size")
    expect(c.sheet.rows.size(2)).toBe(20)
  })

  it("wrapped text fits the column width", () => {
    const { c, select, type } = setup()
    type("A1", "one two three four five six")
    select("A1")
    c.format.apply({ wrap: true }, "Wrap Text")
    // 64 px wide: 58 px inside, 8 characters a line.
    expect(c.sheet.rows.size(0)).toBe(4 * 20)
  })

  it("a row sized by hand keeps its height", () => {
    const { c, select, type } = setup()
    c.setLineSize("rows", [4], 30)
    type("A5", "a\nb\nc")
    expect(c.sheet.rows.size(4)).toBe(30)
    select("A5")
  })
})
