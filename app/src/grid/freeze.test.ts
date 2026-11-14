import { parseRange, selectRange } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { sliderOf, stepZoom, zoomOfSlider } from "@/shell/StatusBar"
import { cellAt, computeLayout, hitTest, rangeRect } from "./layout"

const measure = (text: string) => text.length * 7

function setup() {
  const c = new SheetController()
  c.viewWidth = 700
  c.viewHeight = 380
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  const layout = () => computeLayout(c, 800, 400, measure)
  return { c, select, layout }
}

describe("frozen panes", () => {
  it("Freeze Panes keeps the rows above and columns left of the active cell; nothing on screen moves", () => {
    const { c, select, layout } = setup()
    c.scrollTo(64 * 2, 20 * 10) // C11 at the top left
    select("E14")
    const before = rangeRect(c, layout(), parseRange("E14")!)
    c.freezePanes()
    expect(c.freeze).toEqual({ row: 13, col: 4, top: 10, left: 2 })
    const l = layout()
    expect(rangeRect(c, l, parseRange("E14")!)).toEqual(before)
    // Rows 11–13 and columns C–D are frozen; the scrolling part starts at E14.
    expect(l.panes).toHaveLength(4)
    expect(l.rows.slice(0, 3).map((r) => r.index)).toEqual([10, 11, 12])
    expect(l.rows[3].index).toBe(13)
    expect(l.cols.slice(0, 2).map((r) => r.index)).toEqual([2, 3])
    // Scrolling moves only the scrolling part, and never above the freeze.
    c.scrollTo(0, 0)
    expect(c.scrollY).toBe(20 * 13)
    expect(c.scrollX).toBe(64 * 4)
    c.scrollTo(64 * 10, 20 * 100)
    const after = layout()
    expect(after.rows[0].index).toBe(10)
    expect(after.rows[3].index).toBe(100)
    expect(hitTest(c, after, after.rowHeaderWidth + 10, after.headerHeight + 5)).toEqual({
      zone: "cell",
      row: 10,
      col: 2,
    })
    expect(cellAt(c, after, after.originX + 1, after.originY + 1)).toEqual({ row: 100, col: 10 })
  })

  it("Freeze Top Row and First Column; Unfreeze puts the view back", () => {
    const { c, layout } = setup()
    c.freezeTopRow()
    expect(c.freeze).toEqual({ row: 1, col: 0, top: 0, left: 0 })
    expect(layout().panes).toHaveLength(2)
    c.scrollTo(0, 20 * 50)
    expect(layout().rows[0].index).toBe(0)
    expect(layout().rows[1].index).toBe(50)
    c.unfreeze()
    expect(c.freeze).toBeNull()
    expect(c.scrollY).toBe(0)
    c.freezeFirstColumn()
    expect(c.freeze).toEqual({ row: 0, col: 1, top: 0, left: 0 })
  })

  it("with A1 active, Freeze Panes freezes at the middle of the window (Excel)", () => {
    const { c } = setup()
    c.freezePanes()
    expect(c.freeze?.row).toBe(Math.floor(190 / 20))
    expect(c.freeze?.col).toBe(Math.floor(350 / 64))
  })

  it("Ctrl+Home goes to the first unfrozen cell; moving into the frozen part doesn't scroll", () => {
    const { c, select } = setup()
    select("C3")
    c.freezePanes()
    select("Z500")
    c.keyReady({ key: "Home", primary: true, shift: false, alt: false })
    expect(c.selection.active).toEqual({ row: 2, col: 2 })
    c.scrollTo(64 * 30, 20 * 300)
    const { scrollX, scrollY } = c
    select("A1")
    expect([c.scrollX, c.scrollY]).toEqual([scrollX, scrollY])
  })

  it("freeze and zoom belong to the sheet", () => {
    const { c } = setup()
    c.freezeTopRow()
    c.setZoom(1.5)
    c.sheets.insert(1)
    expect(c.freeze).toBeNull()
    expect(c.zoom).toBe(1)
    c.activate(0)
    expect(c.freeze?.row).toBe(1)
    expect(c.zoom).toBe(1.5)
  })
})

describe("zoom", () => {
  it("stays within 10%–400%", () => {
    const { c } = setup()
    c.setZoom(9)
    expect(c.zoom).toBe(4)
    c.setZoom(0.01)
    expect(c.zoom).toBe(0.1)
  })

  it("the status bar slider has 100% in the middle and steps by tens", () => {
    expect(sliderOf(1)).toBe(500)
    expect(sliderOf(0.1)).toBe(0)
    expect(sliderOf(4)).toBe(1000)
    expect(zoomOfSlider(500)).toBe(1)
    expect(zoomOfSlider(1000)).toBe(4)
    expect(stepZoom(1, 1)).toBe(1.1)
    expect(stepZoom(1.15, -1)).toBe(1.1)
    expect(stepZoom(0.1, -1)).toBe(0.1)
    expect(stepZoom(4, 1)).toBe(4)
  })

  it("Zoom to Selection fits the selection and puts it at the top left", () => {
    const { c, select } = setup()
    select("B2:E11") // 256 × 200 px
    c.zoomToSelection(512, 600)
    expect(c.zoom).toBe(2)
    expect(c.scrollX).toBe(64)
    expect(c.scrollY).toBe(20)
  })
})
