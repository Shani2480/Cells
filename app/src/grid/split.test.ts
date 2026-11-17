import { parseRange, selectRange } from "@cells/core"
import { openXlsx, saveXlsx } from "@cells/io"
import { describe, expect, it } from "vitest"
import { SheetController, SPLIT_MIN } from "@/sheet/controller"
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

describe("split panes", () => {
  it("Split goes above and left of the active cell; nothing on screen moves", () => {
    const { c, select, layout } = setup()
    select("C5")
    const before = rangeRect(c, layout(), parseRange("C5")!)
    c.toggleSplit()
    expect(c.split).toEqual({ width: 128, height: 80, scrollX: 0, scrollY: 0 })
    expect(rangeRect(c, layout(), parseRange("C5")!)).toEqual(before)
    expect(layout().panes).toHaveLength(4)
    c.toggleSplit()
    expect(c.split).toBeNull()
    expect([c.scrollX, c.scrollY]).toEqual([0, 0])
  })

  it("from the top-left cell the split goes through the middle of the window", () => {
    const { c } = setup()
    c.toggleSplit()
    expect(c.split!.width).toBeGreaterThan(300)
    expect(c.split!.height).toBeGreaterThan(150)
  })

  it("from a cell in the first column the split is across only", () => {
    const { c, select } = setup()
    select("A6")
    c.toggleSplit()
    expect(c.split).toMatchObject({ width: 0, height: 100 })
  })

  it("each pane scrolls on its own and can show the same rows", () => {
    const { c, select, layout } = setup()
    select("A6")
    c.toggleSplit()
    c.scrollTo(0, 0)
    const l = layout()
    // The top pane shows rows 1–5, the bottom pane starts at row 1 too.
    expect(l.rows.filter((r) => r.frozen).map((r) => r.index)).toEqual([0, 1, 2, 3, 4])
    expect(l.rows.find((r) => !r.frozen)!.index).toBe(0)
    c.scrollSplitTo(0, 20 * 50)
    const m = layout()
    expect(m.rows.find((r) => r.frozen)!.index).toBe(50)
    // A click in the top pane picks the row it shows there.
    expect(cellAt(c, m, 60, m.headerHeight + 5).row).toBe(50)
    expect(cellAt(c, m, 60, m.originY + 5).row).toBe(0)
  })

  it("the bar is grabbed, moved, and dragged off the edge to remove it", () => {
    const { c, select, layout } = setup()
    select("C5")
    c.toggleSplit()
    const l = layout()
    expect(hitTest(c, l, l.originX, 200)).toEqual({ zone: "split", x: true, y: false })
    expect(hitTest(c, l, l.originX, l.originY)).toEqual({ zone: "split", x: true, y: true })
    c.moveSplit(200, c.split!.height)
    expect(c.split!.width).toBe(200)
    c.moveSplit(SPLIT_MIN - 1, c.split!.height)
    expect(c.split).toMatchObject({ width: 0, height: 80 })
    c.moveSplit(0, 0)
    expect(c.split).toBeNull()
  })

  it("Freeze Panes freezes at the split, and Split on frozen panes splits there", () => {
    const { c, select } = setup()
    select("C5")
    c.freezePanes()
    c.toggleSplit()
    expect(c.freeze).toBeNull()
    expect(c.split).toEqual({ width: 128, height: 80, scrollX: 0, scrollY: 0 })
    c.freezePanes()
    expect(c.split).toBeNull()
    expect(c.freeze).toEqual({ row: 4, col: 2, top: 0, left: 0 })
  })

  it("is saved in Excel files and read back", async () => {
    const { c, select } = setup()
    select("C5")
    c.toggleSplit()
    c.scrollSplitTo(0, 20 * 10)
    const views = c.saveViews()
    expect(views[0].split).toEqual({ width: 128, height: 80, topLeft: { row: 10, col: 0 } })
    const bytes = await saveXlsx(c.wb, { views, now: new Date(Date.UTC(2026, 0, 1)) })
    const back = await openXlsx(bytes, "Split.xlsx")
    expect(back.views[0].split).toEqual(views[0].split)
    const d = new SheetController(back.workbook)
    d.loadViews(back.views, 0)
    expect(d.split).toEqual({ width: 128, height: 80, scrollX: 0, scrollY: 200 })
  })
})
