import { describe, expect, it } from "vitest"
import { OFFICE_THEME, slicerLayout, slicerLook } from "../src"

// Colours sampled from Excel 365's pictures of each style under the Office theme (2023).
describe("slicer styles", () => {
  it("take the theme's accent as Excel does", () => {
    const light1 = slicerLook("SlicerStyleLight1", OFFICE_THEME)
    expect(light1.frame).toBe("#156082")
    expect(light1.selected.fill).toBe("#83CCEB")
    expect(light1.selectedNoData.fill).toBe("#C0E6F5")
    const dark2 = slicerLook("SlicerStyleDark2", OFFICE_THEME)
    expect(dark2.selected).toEqual({ fill: "#E97132", border: "#E97132", text: "#FFFFFF" })
    expect(dark2.selectedNoData.text).toBe("#BE5014")
    expect(slicerLook("SlicerStyleSomethingElse", OFFICE_THEME)).toEqual(light1)
  })

  it("lays buttons out as Excel does: below the header, 3 px apart, scrolled when they don't fit", () => {
    const s = { header: true, columns: 1, buttonHeight: 27 }
    const at = slicerLayout(s, 192, 280, 5)
    expect(at.buttons[0]).toEqual({ x: 9, y: 34, w: 174, h: 27 })
    expect(at.buttons[1].y).toBe(64)
    expect(at.scroll).toBeNull()
    const long = slicerLayout({ ...s, startItem: 3 }, 192, 150, 20)
    expect(long.scroll).toEqual({ total: 20, shown: 3 })
    expect(long.first).toBe(3)
    expect(slicerLayout({ ...s, columns: 2 }, 192, 280, 4).buttons[1].x).toBeGreaterThan(90)
  })
})
