import { StyleTable } from "@cells/core"
import { describe, expect, it } from "vitest"
import { fontFor, formatColor, layoutMarks, lookOf, prefersDarkText, shownText } from "./cell-look"
import { wrapLines } from "./draw-cells"

/** Every character 7 px wide, whatever the font. */
const measure = (text: string) => text.length * 7

const ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

function look(style: Parameters<StyleTable["intern"]>[0]) {
  const styles = new StyleTable()
  return lookOf(styles, styles.intern(style))
}

describe("cell look", () => {
  it("Accounting: the $ sits at the left edge and the amount at the right, with gaps for the brackets", () => {
    const l = look({ numberFormat: ACCOUNTING })
    const shown = shownText(1234.5, l, 120, 3, measure)
    const laid = layoutMarks(shown.text, l.font, measure, 120 - 6)
    expect(laid.runs.map((r) => r.text)).toEqual(["$", "1,234.50"])
    expect(laid.runs[0].x).toBe(7) // after the gap for "("
    expect(laid.runs[1].x + 7 * 8 + 7).toBe(114) // the amount ends one ")" before the right edge
  })

  it("format colours: [Red] and [ColorN]", () => {
    const l = look({ numberFormat: "0;[Red]-0" })
    expect(shownText(-5, l, 64, 3, measure)).toMatchObject({ text: "-5", color: "#FF0000" })
    expect(formatColor("color10")).toBe("#008000")
  })

  it("a formatted number too wide for its column shows ####", () => {
    const l = look({ numberFormat: "#,##0.00" })
    expect(shownText(123456789, l, 40, 3, measure).text).toMatch(/^#+$/)
    expect(shownText(12, l, 64, 3, measure).text).toBe("12.00")
  })

  it("text with a Text format section, and horizontal alignment overrides", () => {
    expect(shownText("abc", look({ numberFormat: '"<"@">"' }), 64, 3, measure).text).toBe("<abc>")
    expect(shownText(5, look({ hAlign: "left" }), 64, 3, measure).align).toBe("left")
    expect(shownText("x", look({ hAlign: "center" }), 64, 3, measure).align).toBe("center")
    expect(shownText(true, look({}), 64, 3, measure).align).toBe("center")
  })

  it("wraps at spaces, breaks long words, keeps line breaks", () => {
    expect(wrapLines("one two three", "", 7 * 8, measure)).toEqual(["one two", "three"])
    expect(wrapLines("abcdefghij", "", 7 * 4, measure)).toEqual(["abcd", "efgh", "ij"])
    expect(wrapLines("a\nb", "", 100, measure)).toEqual(["a", "b"])
  })

  it("rotation follows Excel's encoding", () => {
    expect(look({ rotation: 45 }).rotation).toBe(45)
    expect(look({ rotation: 135 }).rotation).toBe(-45)
    expect(look({ rotation: 255 }).rotation).toBe("stacked")
  })

  it("fonts: size in points, bold and italic, a named font before the fallbacks", () => {
    const { font, px } = fontFor({ fontSize: 12, bold: true, italic: true, fontName: "Georgia" })
    expect(px).toBe(16)
    expect(font.startsWith('italic bold 16px "Georgia", ')).toBe(true)
  })

  it("picks readable automatic text on fills", () => {
    expect(prefersDarkText("#ffff00")).toBe(true)
    expect(prefersDarkText("#002060")).toBe(false)
  })
})
