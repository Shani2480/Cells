import { describe, expect, it } from "vitest"
import { recordText } from "./text-layer"

/** A canvas context as far as the recorder uses it: text 6 units a character wide, a font box 9 high above "top". */
function painter(transform = { a: 2, b: 0, c: 0, d: 2, e: 10, f: 20 }) {
  const drawn: string[] = []
  const p = {
    font: "12px Aptos",
    textAlign: "left" as CanvasTextAlign,
    textBaseline: "alphabetic" as CanvasTextBaseline,
    direction: "ltr" as CanvasDirection,
    fillText(text: string) {
      drawn.push(text)
    },
    measureText(text: string) {
      // The font's top is 9 above the alphabetic baseline, 0 above "top".
      return { width: text.length * 6, fontBoundingBoxAscent: p.textBaseline === "top" ? 0 : 9 }
    },
    getTransform() {
      const { a, b, c, d, e, f } = transform
      return {
        a,
        b,
        c,
        d,
        e,
        f,
        transformPoint: (pt: { x: number; y: number }) => ({ x: a * pt.x + c * pt.y + e, y: b * pt.x + d * pt.y + f }),
      }
    },
  }
  return { p: p as unknown as CanvasRenderingContext2D, drawn }
}

const page = { width: 600, height: 800 }

describe("PDF text layer: the text a page draws", () => {
  it("records each piece in points, where its baseline starts, still drawing it", () => {
    const { p, drawn } = painter()
    const stop = recordText(p, 2, page)
    p.fillText("Region", 5, 30)
    p.textAlign = "right"
    p.fillText("1,250.00", 100, 30)
    p.textAlign = "center"
    p.fillText("  ", 50, 30)
    const text = stop()
    expect(drawn).toEqual(["Region", "1,250.00", "  "])
    // 2 canvas pixels a point and a transform of 2: a canvas unit is a point, moved by (5, 10) points.
    expect(text).toEqual([
      { text: "Region", x: 10, y: 40, size: 12, width: 36 },
      { text: "1,250.00", x: 57, y: 40, size: 12, width: 48 },
    ])
  })

  it("finds the alphabetic baseline of text set from its top, and keeps a turned piece's direction", () => {
    const { p } = painter({ a: 0, b: -2, c: 2, d: 0, e: 100, f: 400 })
    const stop = recordText(p, 2, page)
    p.textBaseline = "top"
    p.fillText("Up", 0, 0)
    const [up] = stop()
    expect(up.angle).toBe(-90)
    // The baseline 9 units under the top, along the turned y axis (which runs to the right on the page).
    expect(up.x).toBeCloseTo(59)
    expect(up.y).toBeCloseTo(200)
    expect(p.textBaseline).toBe("top")
  })

  it("leaves out text drawn off the page, honours maxWidth, and puts fillText back", () => {
    const { p, drawn } = painter({ a: 2, b: 0, c: 0, d: 2, e: 0, f: 0 })
    const own = p.fillText
    const stop = recordText(p, 2, page)
    p.fillText("Far away", 5000, 10)
    p.fillText("Squeezed", 0, 10, 24)
    const text = stop()
    expect(text.map((t) => [t.text, t.width])).toEqual([["Squeezed", 24]])
    expect(p.fillText).toBe(own)
    p.fillText("After", 0, 0)
    expect(drawn.at(-1)).toBe("After")
  })
})
