import { describe, expect, it } from "vitest"
import { colorCss, defaultShape, defaultTextBox, OFFICE_THEME } from "../src"

describe("new shapes look as in Excel 365", () => {
  it("a shape: accent 1 with a 15% shade of it as a 1.5 pt outline; a text box: white with a #BCBCBC outline", () => {
    const shape = defaultShape("rect")
    expect(colorCss((shape.fill as { color: never }).color, OFFICE_THEME).hex).toBe("#156082")
    expect(colorCss(shape.line!.color, OFFICE_THEME).hex).toBe("#042433")
    expect(shape.line!.width).toBe(1.5)
    const box = defaultTextBox()
    expect(colorCss(box.line!.color, OFFICE_THEME).hex).toBe("#BCBCBC")
    expect(box.line!.width).toBe(0.75)
  })
})
