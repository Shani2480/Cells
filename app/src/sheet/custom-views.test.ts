import { parseRange, selectRange } from "@cells/core"
import { openXlsx, saveXlsx } from "@cells/io"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"

function setup() {
  const c = new SheetController()
  c.viewWidth = 1000
  c.viewHeight = 600
  c.sheets.insert(1)
  c.activate(0)
  return c
}

describe("Custom Views", () => {
  it("Show puts back the selection, zoom, active sheet and hidden rows a view was added with", () => {
    const c = setup()
    c.setSelection(selectRange(parseRange("C5")!))
    c.setZoom(1.25)
    c.customViews.add("Full", true, true)
    c.setSelection(selectRange(parseRange("B2:B3")!))
    c.hideSelected("rows", true)
    c.setZoom(0.5)
    c.activate(1)
    c.customViews.show("full")
    expect(c.sheetIndex).toBe(0)
    expect(c.zoom).toBe(1.25)
    expect(c.selection.active).toEqual({ row: 4, col: 2 })
    expect(c.sheet.rows.isHidden(1)).toBe(false)
    // Showing it changed the rows as one undo step.
    c.undo()
    expect(c.wb.sheets[0].rows.isHidden(1)).toBe(true)
  })

  it("a view without hidden rows or print settings leaves them as they are", () => {
    const c = setup()
    c.customViews.add("Plain", false, false)
    c.setSelection(selectRange(parseRange("A3")!))
    c.hideSelected("rows", true)
    c.print.setOrientation("landscape")
    c.customViews.show("Plain")
    expect(c.sheet.rows.isHidden(2)).toBe(true)
    expect(c.sheet.pageSetup.orientation).toBe("landscape")
  })

  it("print settings come back with a view that has them", () => {
    const c = setup()
    c.customViews.add("Print", true, false)
    c.print.setOrientation("landscape")
    c.customViews.show("Print")
    expect(c.sheet.pageSetup.orientation).toBe("portrait")
  })

  it("adding a name again replaces the view; Delete removes it", () => {
    const c = setup()
    c.customViews.add("A", true, true)
    const guid = c.wb.customViews[0].guid
    c.customViews.add("a", false, true)
    expect(c.wb.customViews).toHaveLength(1)
    expect(c.wb.customViews[0].guid).toBe(guid)
    expect(c.wb.customViews[0].printSettings).toBe(false)
    c.customViews.remove("A")
    expect(c.wb.customViews).toHaveLength(0)
  })

  it("views are saved in Excel files and work after opening", async () => {
    const c = setup()
    c.setSelection(selectRange(parseRange("D4")!))
    c.hideSelected("cols", true)
    c.customViews.add("No D", true, true)
    c.setSelection(selectRange(parseRange("D1")!))
    c.hideSelected("cols", false)
    const bytes = await saveXlsx(c.wb, { views: c.saveViews(), now: new Date(Date.UTC(2026, 0, 1)) })
    const back = await openXlsx(bytes, "Views.xlsx")
    const d = new SheetController(back.workbook)
    d.loadViews(back.views, 0)
    d.customViews.show("No D")
    expect(d.sheet.cols.isHidden(3)).toBe(true)
    expect(d.selection.active).toEqual({ row: 3, col: 3 })
  })
})
