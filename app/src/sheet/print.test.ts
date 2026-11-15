import { describe, expect, it, vi } from "vitest"
import { MARGIN_PRESETS, writeInputOp } from "@cells/core"
import { headerFooterLines, type PrintContext } from "@/print/render"
import { SheetController } from "./controller"
import { marginPreset } from "./print"

function filled(rows: number, cols: number) {
  const c = new SheetController()
  const ops = []
  for (let r = 0; r < rows; r++)
    for (let col = 0; col < cols; col++) ops.push(writeInputOp(0, c.sheet, c.wb.styles, r, col, String(r * cols + col)))
  c.wb.transact(ops, { label: "Setup" })
  c.shell.alert = vi.fn()
  c.print.defaultPaper = 1
  return c
}

describe("printing", () => {
  it("changes the page setup as one undo step and shows the page breaks", () => {
    const c = filled(3, 3)
    expect(c.print.showBreaks).toBe(false)
    c.print.setOrientation("landscape")
    c.print.setMargins(MARGIN_PRESETS.narrow)
    expect(c.print.setup().orientation).toBe("landscape")
    expect(marginPreset(c.print.setup().margins)).toBe("narrow")
    expect(c.print.showBreaks).toBe(true)
    c.undo()
    expect(marginPreset(c.print.setup().margins)).toBe("normal")
    c.undo()
    expect(c.print.setup().orientation).toBe("portrait")
  })

  it("sets, adds to and clears the print area", () => {
    const c = filled(3, 3)
    c.goTo({ r0: 0, c0: 0, r1: 1, c1: 1 })
    expect(c.print.canAddToPrintArea()).toBe(false)
    c.print.setPrintArea()
    expect(c.print.setup().printArea).toEqual([{ r0: 0, c0: 0, r1: 1, c1: 1 }])
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    expect(c.print.canAddToPrintArea()).toBe(false)
    c.goTo({ r0: 5, c0: 0, r1: 6, c1: 0 })
    expect(c.print.canAddToPrintArea()).toBe(true)
    c.print.addToPrintArea()
    expect(c.print.setup().printArea).toHaveLength(2)
    expect(c.print.pages().pages).toHaveLength(2)
    c.print.clearPrintArea()
    expect(c.print.setup().printArea).toBeNull()
  })

  it("inserts page breaks above and left of the cell, and removes them", () => {
    const c = filled(60, 4)
    c.goTo({ r0: 20, c0: 2, r1: 20, c1: 2 })
    c.print.insertBreak()
    expect(c.print.setup().rowBreaks).toEqual([20])
    expect(c.print.setup().colBreaks).toEqual([2])
    expect(c.print.breaks().rows).toContain(20)
    expect(c.print.pages().pages.length).toBeGreaterThanOrEqual(4)
    expect(c.print.hasBreakAtCell()).toBe(true)
    // Remove takes the breaks along the cell's top and left edges.
    c.print.removeBreak()
    expect([c.print.setup().rowBreaks, c.print.setup().colBreaks]).toEqual([[], []])
    // In column A only a row break goes in.
    c.goTo({ r0: 10, c0: 0, r1: 10, c1: 0 })
    c.print.insertBreak()
    expect([c.print.setup().rowBreaks, c.print.setup().colBreaks]).toEqual([[10], []])
    c.print.resetBreaks()
    expect([c.print.setup().rowBreaks, c.print.setup().colBreaks]).toEqual([[], []])
  })

  it("fits to pages, and the scale only takes 10 to 400 while not fitting", () => {
    const c = filled(200, 4)
    expect(c.print.pages().pages.length).toBeGreaterThan(1)
    c.print.setFit("height", 1)
    expect(c.print.setup().fitTo).toEqual({ width: 0, height: 1 })
    expect(c.print.pages().pages).toHaveLength(1)
    c.print.setFit("height", 0)
    expect(c.print.setup().fitTo).toBeNull()
    expect(c.print.setScale(5)).toBe(false)
    expect(c.print.setScale(50)).toBe(true)
    expect(c.print.setup().scale).toBe(50)
  })

  it("toggles printed gridlines and headings", () => {
    const c = filled(2, 2)
    c.print.toggleGridlines()
    c.print.toggleHeadings()
    expect([c.print.setup().gridlines, c.print.setup().headings]).toEqual([true, true])
  })
})

describe("header and footer codes", () => {
  const c = new SheetController()
  const page = c.print.pages().pages[0] ?? {
    area: 0,
    rows: [0, 0],
    cols: [0, 0],
    titleRows: null,
    titleCols: null,
    number: 3,
  }
  const ctx: PrintContext = {
    fileName: "Budget.xlsx",
    folder: "/home/alex/Documents",
    now: new Date(2026, 0, 2, 9, 5),
    total: 7,
  }

  it("reads page numbers, the page count, file, folder and sheet", () => {
    const p = { ...page, number: 3 }
    expect(headerFooterLines("Page &P of &N", p, c.sheet, ctx)).toEqual(["Page 3 of 7"])
    expect(headerFooterLines("&P+2 / &P-1", p, c.sheet, ctx)).toEqual(["5 / 2"])
    expect(headerFooterLines("&F in &Z, &A", p, c.sheet, ctx)).toEqual(["Budget.xlsx in /home/alex/Documents, Sheet1"])
    expect(headerFooterLines("A && B", p, c.sheet, ctx)).toEqual(["A & B"])
  })

  it("leaves out fonts, sizes and styles, and splits lines", () => {
    const p = { ...page, number: 1 }
    expect(headerFooterLines('&"Arial,Bold"&14&BTitle&B\nline two', p, c.sheet, ctx)).toEqual(["Title", "line two"])
    expect(headerFooterLines("&KFF0000Red", p, c.sheet, ctx)).toEqual(["Red"])
    expect(headerFooterLines("&D", p, c.sheet, ctx)).toEqual([ctx.now.toLocaleDateString()])
  })
})
