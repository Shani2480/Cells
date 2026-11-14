import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { drawGrid, type Painter } from "./draw"
import { computeLayout, hitTest } from "./layout"
import { LIGHT_THEME } from "./theme"

/** A painter that records text and ignores everything else. */
function recorder() {
  const texts: { text: string; x: number; y: number; align: string }[] = []
  const noop = () => {}
  const painter = {
    save: noop,
    restore: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    fill: noop,
    fillRect: noop,
    stroke: noop,
    strokeRect: noop,
    moveTo: noop,
    lineTo: noop,
    closePath: noop,
    setLineDash: noop,
    setTransform: noop,
    clearRect: noop,
    fillStyle: "",
    strokeStyle: "",
    lineWidth: 1,
    font: "",
    textAlign: "left" as CanvasTextAlign,
    textBaseline: "alphabetic" as CanvasTextBaseline,
    lineDashOffset: 0,
    fillText(text: string, x: number, y: number) {
      texts.push({ text, x, y, align: this.textAlign })
    },
  } satisfies Painter
  return { painter, texts }
}

const measure = (text: string) => text.length * 7

function draw(c: SheetController) {
  const layout = computeLayout(c, 800, 400, measure)
  const { painter, texts } = recorder()
  drawGrid({ painter, layout, controller: c, theme: LIGHT_THEME, dpr: 1, measure, antsOffset: 0 })
  return { layout, texts: texts.map((t) => t.text), raw: texts }
}

describe("drawGrid", () => {
  it("draws headers, numbers right-aligned and text left-aligned", () => {
    const c = new SheetController()
    c.wb.transact(
      [
        {
          kind: "setCells",
          sheet: 0,
          cells: [
            { row: 0, col: 0, data: { value: "Name" } },
            { row: 0, col: 1, data: { value: 12.5 } },
          ],
        },
      ],
      { label: "Set" },
    )
    const { texts, raw } = draw(c)
    expect(texts).toContain("A")
    expect(texts).toContain("1")
    expect(texts).toContain("Name")
    expect(raw.find((t) => t.text === "12.5")?.align).toBe("right")
    expect(raw.find((t) => t.text === "Name")?.align).toBe("left")
  })

  it("shows #### for a number too wide for its column", () => {
    const c = new SheetController()
    c.sheet.cols.setSize(0, 20)
    c.wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: { value: 123456789 } }] }], {
      label: "Set",
    })
    expect(draw(c).texts.some((t) => /^#+$/.test(t))).toBe(true)
  })

  it("skips hidden rows and widens the row header for large row numbers", () => {
    const c = new SheetController()
    c.sheet.rows.setHidden(1, true)
    const first = draw(c)
    expect(first.texts).not.toContain("2")
    expect(first.layout.rows.map((r) => r.index).slice(0, 2)).toEqual([0, 2])
    c.scrollTo(0, c.sheet.rows.offset(999_990))
    const far = draw(c)
    expect(far.layout.rowHeaderWidth).toBeGreaterThan(first.layout.rowHeaderWidth)
    expect(far.texts).toContain("999991")
  })

  it("hit-tests headers, resize edges and the fill handle", () => {
    const c = new SheetController()
    const { layout } = draw(c)
    const x = layout.rowHeaderWidth
    expect(hitTest(c, layout, 2, 2).zone).toBe("corner")
    expect(hitTest(c, layout, x + 30, 5)).toEqual({ zone: "colHeader", col: 0 })
    expect(hitTest(c, layout, x + 64, 5)).toEqual({ zone: "colResize", col: 0 })
    expect(hitTest(c, layout, 5, layout.headerHeight + 30)).toEqual({ zone: "rowHeader", row: 1 })
    expect(hitTest(c, layout, x + 64, layout.headerHeight + 20).zone).toBe("fillHandle")
    expect(hitTest(c, layout, x + 100, layout.headerHeight + 50)).toEqual({ zone: "cell", row: 2, col: 1 })
  })
})
