import { selectCell } from "@cells/core"
import { openXlsx, saveXlsx } from "@cells/io"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { PAGE_GAP } from "@/sheet/page-map"
import { cellAt, colX, computeLayout, rowY } from "./layout"
import { pagePartAt, visiblePages } from "./page-layout"
import { rulerUnits } from "./Ruler"

const measure = (text: string) => text.length * 7

/** 120 rows of numbers in A:C: more than one printed page. */
function setup() {
  const c = new SheetController()
  c.viewWidth = 1000
  c.viewHeight = 600
  const cells = []
  for (let row = 0; row < 120; row++) for (let col = 0; col < 3; col++) cells.push({ row, col, data: { value: row } })
  c.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "x", before: c.selection })
  return c
}

describe("Page Layout view", () => {
  it("puts the rows on pages as printing does, inside the margins, with a gap between pages", () => {
    const c = setup()
    c.setView("pageLayout")
    const map = c.pageMap()!
    const printed = c.print.current().pages
    // The printed pages start where Page Layout's pages start.
    expect(map.rows.bands.slice(0, printed.length).map((b) => b.first)).toEqual(printed.map((p) => p.rows[0]))
    // A1 sits inside the first page's top and left margins.
    expect(map.rows.pos(0)).toBeCloseTo(PAGE_GAP + map.rows.margin)
    expect(map.cols.pos(0)).toBeCloseTo(PAGE_GAP + map.cols.margin)
    // The second page's first row starts a page further on.
    const second = map.rows.bands[1]
    expect(map.rows.pos(second.first)).toBeCloseTo(second.start + map.rows.margin)
    expect(second.start - map.rows.bands[0].start).toBeCloseTo(map.rows.page + PAGE_GAP)
    // The row at a position on the second page is that page's row.
    expect(map.rows.indexAt(map.rows.pos(second.first) + 1)).toBe(second.first)
    // Blank pages follow the cells in use.
    expect(map.rows.bands.length).toBeGreaterThan(printed.length)
  })

  it("fills a page with the columns that fit, past the last one in use, unless a print area ends it", () => {
    const c = setup()
    c.setView("pageLayout")
    // A:C print on one page; the page carries on with D, E… as Excel's Page Layout shows it.
    const first = c.pageMap()!.cols.bands[0]
    expect(first.first).toBe(0)
    expect(first.last).toBeGreaterThan(5)
    c.print.change((s) => (s.printArea = [{ r0: 0, c0: 0, r1: 119, c1: 2 }]))
    expect(c.pageMap()!.cols.bands[0].last).toBe(2)
  })

  it("an empty sheet still shows whole pages", () => {
    const c = new SheetController()
    c.setView("pageLayout")
    const map = c.pageMap()!
    expect(map.rows.bands[0].last).toBeGreaterThan(30)
    expect(map.cols.bands[0].last).toBeGreaterThan(5)
  })

  it("the layout places cells on their pages, and clicks in a margin find no cell", () => {
    const c = setup()
    c.setView("pageLayout")
    const layout = computeLayout(c, 1000, 600, measure)
    expect(layout.pages).toBeDefined()
    const sheet = c.sheet
    const x = colX(sheet, layout, 0)
    const y = rowY(sheet, layout, 0)
    expect(cellAt(c, layout, x + 2, y + 2)).toEqual({ row: 0, col: 0 })
    const page = visiblePages(layout)[0]
    expect(pagePartAt(layout, x + 2, y + 2)).toBeNull()
    expect(pagePartAt(layout, page.x + page.w / 2, page.y + 4)).toBe("header")
    expect(pagePartAt(layout, page.x + 2, y + 20)).toBe("margin")
    expect(pagePartAt(layout, layout.rowHeaderWidth + 2, layout.headerHeight + 2)).toBe("desk")
    // Opening at A1 shows the first page whole, from its top-left corner.
    expect(c.scrollX).toBe(0)
    expect(c.scrollY).toBe(0)
  })

  it("keeps the same cell at the top left when switching views, and takes frozen panes away", () => {
    const c = setup()
    c.scrollTo(0, c.sheet.rows.offset(60))
    c.setSelection(selectCell({ row: 62, col: 1 }))
    c.freezeTopRow()
    c.setView("pageLayout")
    expect(c.freeze).toBeNull()
    expect(c.topLeft.row).toBe(60)
    c.setView("normal")
    expect(c.topLeft.row).toBe(60)
  })

  it("scrolls a cell on a later page into view", () => {
    const c = setup()
    c.setView("pageLayout")
    c.scrollIntoView({ row: 110, col: 0 })
    const layout = computeLayout(c, 1000, 600, measure)
    const y = rowY(c.sheet, layout, 110)
    expect(y).toBeGreaterThanOrEqual(layout.headerHeight)
    expect(y + c.sheet.rows.size(110)).toBeLessThanOrEqual(layout.headerHeight + c.paneHeight + 1)
  })

  it("is saved as the sheet's view, with where it was scrolled", async () => {
    const c = setup()
    c.setView("pageLayout")
    const second = c.pageMap()!.rows.bands[1].first
    c.scrollTo(0, c.pageMap()!.rows.pos(second))
    const bytes = await saveXlsx(c.wb, { views: c.saveViews(), now: new Date(Date.UTC(2026, 0, 1)) })
    const back = await openXlsx(bytes, "Pages.xlsx")
    expect(back.views[0].mode).toBe("pageLayout")
    const d = new SheetController(back.workbook)
    d.viewWidth = 1000
    d.viewHeight = 600
    d.loadViews(back.views, 0)
    expect(d.view).toBe("pageLayout")
    expect(d.topLeft.row).toBe(second)
  })

  it("measures the ruler in inches in the US and centimetres elsewhere", () => {
    expect(rulerUnits("en-US")).toBe("in")
    expect(rulerUnits("en-GB")).toBe("cm")
    expect(rulerUnits("nl-NL")).toBe("cm")
  })
})
