import { describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"

function book() {
  const c = new SheetController()
  const rows = [
    [3, 5, 2, 8],
    [1, -4, 9, 2],
    [5, 5, 6, 4],
  ]
  rows.forEach((row, r) => row.forEach((v, col) => c.sheet.cells.set(r, col, { value: v })))
  c.shell.alert = vi.fn()
  return c
}
const at = (c: SheetController, row: number, col: number) => c.goTo({ r0: row, c0: col, r1: row, c1: col })

describe("Insert › Sparklines and the Sparkline tab", () => {
  it("creates one sparkline per location cell, each on its row of data, as one undo step", () => {
    const c = book()
    expect(c.sparklines.create("line", "A1:D3", "E1:E3")).toBe(true)
    const g = c.sheet.sparklines[0]
    expect(g.sparklines.map((s) => s.source)).toEqual(["Sheet1!A1:D1", "Sheet1!A2:D2", "Sheet1!A3:D3"])
    expect(c.selection.active).toEqual({ row: 0, col: 4 })
    expect(c.sparklines.active()).toBe(g)
    c.undo()
    expect(c.sheet.sparklines).toEqual([])
  })

  it("shows Excel's messages for ranges that don't fit", () => {
    const c = book()
    expect(c.sparklines.create("line", "A1:D3", "E1:F3")).toBe(false)
    expect(c.shell.alert).toHaveBeenCalledWith(expect.stringMatching(/same column or same row/))
    expect(c.sparklines.create("line", "nonsense", "E1")).toBe(false)
    expect(c.sparklines.create("line", "A1:D3", "E1:E2")).toBe(false)
  })

  it("changes type, shows points, style, colours and axis for the selected groups", () => {
    const c = book()
    c.sparklines.create("line", "A1:D3", "E1:E3")
    at(c, 1, 4)
    c.sparklines.setType("column")
    c.sparklines.toggleFlag("high")
    c.sparklines.setStyle(30)
    c.sparklines.setColor("negative", "#00FF00")
    c.sparklines.setScale("min", "custom", -2)
    c.sparklines.setEmpty("zero", true)
    const g = c.sheet.sparklines[0]
    expect(g).toMatchObject({
      type: "column",
      high: true,
      emptyAs: "zero",
      showHidden: true,
      minAxis: "custom",
      manualMin: -2,
    })
    expect(g.colors.series).toBe("#5F5F5F")
    expect(g.colors.negative).toBe("#00FF00")
    expect(c.sparklines.styleIndex()).toBe(-1)
    expect(c.sparklines.setDateRange("A1:D1")).toBe(true)
    expect(c.sheet.sparklines[0].dateRange).toBe("Sheet1!A1:D1")
  })

  it("ungroups, groups again and clears", () => {
    const c = book()
    c.sparklines.create("line", "A1:D3", "E1:E3")
    c.goTo({ r0: 0, c0: 4, r1: 2, c1: 4 })
    c.sparklines.ungroup()
    expect(c.sheet.sparklines).toHaveLength(3)
    expect(c.sparklines.canGroup()).toBe(true)
    c.sparklines.group()
    expect(c.sheet.sparklines).toHaveLength(1)
    c.goTo({ r0: 1, c0: 4, r1: 1, c1: 4 })
    c.sparklines.clear()
    expect(c.sheet.sparklines[0].sparklines.map((s) => s.row)).toEqual([0, 2])
    at(c, 0, 4)
    c.sparklines.clear(true)
    expect(c.sheet.sparklines).toEqual([])
  })

  it("edits one sparkline's data, and a new group takes over cells from an old one", () => {
    const c = book()
    c.sparklines.create("line", "A1:D3", "E1:E3")
    at(c, 0, 4)
    expect(c.sparklines.editSingle("B1:D1")).toBe(true)
    expect(c.sheet.sparklines[0].sparklines[0].source).toBe("Sheet1!B1:D1")
    c.sparklines.create("column", "A1:D1", "E1")
    expect(c.sheet.sparklines.map((g) => [g.type, g.sparklines.length])).toEqual([
      ["line", 2],
      ["column", 1],
    ])
  })
})
