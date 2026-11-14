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
    arc: noop,
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
    translate: noop,
    rotate: noop,
    createLinearGradient: () => ({ addColorStop: noop }) as unknown as CanvasGradient,
    fillStyle: "" as string | CanvasGradient,
    strokeStyle: "",
    lineWidth: 1,
    font: "",
    textAlign: "left" as CanvasTextAlign,
    textBaseline: "alphabetic" as CanvasTextBaseline,
    lineDashOffset: 0,
    globalAlpha: 1,
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
    const { texts, raw, layout } = draw(c)
    expect(texts).toContain("A")
    expect(texts).toContain("1")
    expect(texts).toContain("Name")
    const [a, b] = layout.cols
    // Numbers end 3 px inside the cell's right edge; text starts 3 px inside its left edge.
    const number = raw.find((t) => t.text === "12.5")!
    expect(number.x + measure("12.5")).toBe(b.start + b.size - 3)
    expect(raw.find((t) => t.text === "Name")!.x).toBe(a.start + 3)
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

  it("shows touch grippers instead of the fill handle after a finger", () => {
    const c = new SheetController()
    c.pointerCell({ row: 1, col: 1 }, { shift: false, ctrl: false })
    c.dragTo({ row: 3, col: 2 })
    c.touchHandles = true
    const { layout } = draw(c)
    const x0 = layout.rowHeaderWidth + 64
    const y0 = layout.headerHeight + 20
    // B2:C4 → top-left at (B, row 2), bottom-right at the end of C4; a finger's width around each counts.
    expect(hitTest(c, layout, x0 + 3, y0 - 4)).toEqual({ zone: "gripper", end: "start" })
    expect(hitTest(c, layout, x0 + 128 + 10, y0 + 60 + 8)).toEqual({ zone: "gripper", end: "end" })
    expect(hitTest(c, layout, x0 + 64, y0 + 30).zone).toBe("cell")
    c.touchHandles = false
    expect(hitTest(c, layout, x0 + 128, y0 + 60).zone).toBe("fillHandle")
  })

  it("draws a frozen-pane frame of a 100,000-row sheet within the frame budget", () => {
    // Phase 3 budget: frozen panes keep the frame budget. This times the layout and drawing code (the canvas
    // itself is a recorder here, so the GPU's share is not in it): best of 5, under half a 60 Hz frame.
    const c = new SheetController()
    const cells = []
    for (let r = 0; r < 100_000; r++)
      for (let col = 0; col < 20; col++) cells.push({ row: r, col, data: { value: r * 20 + col } })
    c.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" })
    c.pointerCell({ row: 4, col: 2 }, { shift: false, ctrl: false })
    c.freezePanes()
    c.scrollTo(c.sheet.cols.offset(5), c.sheet.rows.offset(50_000))
    let ms = Infinity
    for (let i = 0; i < 5; i++) {
      const start = performance.now()
      const layout = computeLayout(c, 1920, 1080, measure)
      drawGrid({
        painter: recorder().painter,
        layout,
        controller: c,
        theme: LIGHT_THEME,
        dpr: 1,
        measure,
        antsOffset: 0,
      })
      ms = Math.min(ms, performance.now() - start)
    }
    console.log(`frozen-pane frame (1920×1080)            ${ms.toFixed(1)} ms`)
    expect(c.freeze).not.toBeNull()
    expect(ms).toBeLessThan(8)
  })
})
