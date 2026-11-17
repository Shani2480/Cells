import { openXlsx, saveXlsx } from "@cells/io"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { breakLineAt, breakLines } from "./break-preview"
import { computeLayout } from "./layout"

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

describe("Page Break Preview", () => {
  it("has its own zoom (60% at first) and gives Normal its zoom back", () => {
    const c = setup()
    c.setZoom(1.2)
    c.setView("pageBreakPreview")
    expect(c.zoom).toBe(0.6)
    c.setZoom(0.8)
    c.setView("normal")
    expect(c.zoom).toBe(1.2)
    c.setView("pageBreakPreview")
    expect(c.zoom).toBe(0.8)
  })

  it("lists the print area's edges and the automatic break between pages", () => {
    const c = setup()
    const { lines, pages } = breakLines(c)
    expect(pages.length).toBeGreaterThan(1)
    const breaks = lines.filter((l) => l.kind === "break" && l.axis === "rows")
    expect(breaks.length).toBe(pages.length - 1)
    expect(breaks.every((l) => !l.manual)).toBe(true)
    expect(lines.filter((l) => l.kind === "edge")).toHaveLength(4)
  })

  it("dragging a break moves it and makes it a manual one; dragged off the area it goes away", () => {
    const c = setup()
    const auto = breakLines(c).lines.find((l) => l.kind === "break")!
    c.print.moveLine(auto, 20)
    expect(c.sheet.pageSetup.rowBreaks).toEqual([20])
    const moved = breakLines(c).lines.find((l) => l.kind === "break" && l.index === 20)!
    expect(moved.manual).toBe(true)
    c.print.moveLine(moved, 500)
    expect(c.sheet.pageSetup.rowBreaks).toEqual([])
  })

  it("dragging the bottom edge sets the print area", () => {
    const c = setup()
    const bottom = breakLines(c).lines.find((l) => l.kind === "edge" && l.axis === "rows" && l.index === 120)!
    c.print.moveLine(bottom, 30)
    expect(c.sheet.pageSetup.printArea).toEqual([{ r0: 0, c0: 0, r1: 29, c1: 2 }])
  })

  it("a press on a blue line grabs it, only in Page Break Preview", () => {
    const c = setup()
    const l = computeLayout(c, 1000, 600, measure)
    expect(breakLineAt(c, l, l.rowHeaderWidth + 1, l.headerHeight + 100)).toBeNull()
    c.setView("pageBreakPreview")
    const m = computeLayout(c, 1000, 600, measure)
    expect(breakLineAt(c, m, m.rowHeaderWidth + 1, m.headerHeight + 50)).toMatchObject({ kind: "edge", axis: "cols" })
  })

  it("the view and each view's zoom are saved in Excel files", async () => {
    const c = setup()
    c.setZoom(1.5)
    c.setView("pageBreakPreview")
    c.setZoom(0.7)
    const views = c.saveViews()
    const bytes = await saveXlsx(c.wb, { views, now: new Date(Date.UTC(2026, 0, 1)) })
    const back = await openXlsx(bytes, "Pages.xlsx")
    expect(back.views[0]).toMatchObject({ mode: "pageBreakPreview", zoom: 0.7 })
    const d = new SheetController(back.workbook)
    d.loadViews(back.views, 0)
    expect(d.view).toBe("pageBreakPreview")
    d.setView("normal")
    expect(d.zoom).toBe(1.5)
  })
})
