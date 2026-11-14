import { describe, expect, it, vi } from "vitest"
import { writeInputOp } from "@cells/core"
import { SheetController } from "./controller"
import { CHART_HEIGHT, CHART_WIDTH, chartBox } from "./charts"

function sheet(rows: string[][]) {
  const c = new SheetController()
  c.wb.transact(
    rows.flatMap((r, row) => r.map((t, col) => (t ? [writeInputOp(0, c.sheet, c.wb.styles, row, col, t)] : [])).flat()),
    { label: "Setup" },
  )
  c.shell.alert = vi.fn()
  return c
}

const SALES = [
  ["", "North", "South"],
  ["Q1", "12", "8"],
  ["Q2", "18", "11"],
  ["Q3", "15", "9"],
]

describe("charts", () => {
  it("Insert charts the data around the cell, selects the chart and opens Chart Design", () => {
    const c = sheet(SALES)
    c.shell.showTab = vi.fn()
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    const chart = c.charts.insert("column")!
    expect(chart.refs!.series.map((s) => s.values)).toEqual(["Sheet1!$B$2:$B$4", "Sheet1!$C$2:$C$4"])
    expect(c.charts.selected).toBe(chart.id)
    expect(c.shell.showTab).toHaveBeenCalledWith("TabChartDesign")
    const box = chartBox(c.sheet, chart.anchor)
    expect([Math.round(box.w), Math.round(box.h)]).toEqual([CHART_WIDTH, CHART_HEIGHT])
    const spec = c.charts.spec()!
    expect(spec.title).toEqual({ show: true, text: "Chart Title" })
    expect(spec.series.map((s) => [s.name, s.values])).toEqual([
      ["North", [12, 18, 15]],
      ["South", [8, 11, 9]],
    ])
    // The look is kept without the numbers: they come from the cells.
    expect((c.sheet.charts[0].spec.series as unknown[]).length).toBe(0)
    c.wb.undo()
    expect(c.sheet.charts).toEqual([])
  })

  it("a one-series chart is titled with the series' name and has no legend", () => {
    const c = sheet(SALES)
    c.goTo({ r0: 0, c0: 0, r1: 3, c1: 1 })
    c.charts.insert("column")
    const spec = c.charts.spec()!
    expect(spec.title.text).toBe("North")
    expect(spec.legend).toBe("none")
  })

  it("draws what the cells hold now", () => {
    const c = sheet(SALES)
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.charts.insert()
    c.wb.transact([writeInputOp(0, c.sheet, c.wb.styles, 1, 1, "40")], { label: "Typing" })
    expect(c.charts.spec()!.series[0].values).toEqual([40, 18, 15])
  })

  it("Delete removes the selected chart; Escape goes back to the cells", () => {
    const c = sheet(SALES)
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.charts.insert()
    expect(c.keyReady({ key: "Escape", primary: false, shift: false, alt: false })).toBe(true)
    expect(c.charts.selected).toBe(null)
    c.charts.select(c.sheet.charts[0].id)
    expect(c.keyReady({ key: "Delete", primary: false, shift: false, alt: false })).toBe(true)
    expect(c.sheet.charts).toEqual([])
    expect(c.charts.selected).toBe(null)
  })

  it("moves and sizes in one undo step each", () => {
    const c = sheet(SALES)
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    const chart = c.charts.insert()!
    c.charts.place(chart.id, { x: 100, y: 200, w: 300, h: 150 })
    const box = chartBox(c.sheet, c.sheet.charts[0].anchor)
    expect([box.x, box.y, box.w, box.h].map(Math.round)).toEqual([100, 200, 300, 150])
    c.wb.undo()
    expect(chartBox(c.sheet, c.sheet.charts[0].anchor).w).toBeCloseTo(CHART_WIDTH)
  })

  it("Change Chart Type, Switch Row/Column and Select Data", () => {
    const c = sheet(SALES)
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.charts.insert()
    c.charts.setVariant("line")
    expect(c.charts.spec()!.type).toBe("line")
    expect(c.charts.dataRange()).toBe("=Sheet1!$A$1:$C$4")
    c.charts.switchRowColumn()
    expect(c.charts.spec()!.series.map((s) => s.name)).toEqual(["Q1", "Q2", "Q3"])
    expect(c.charts.spec()!.categories).toEqual(["North", "South"])
    c.charts.removeSeries(0)
    expect(c.charts.spec()!.series.map((s) => s.name)).toEqual(["Q2", "Q3"])
    expect(c.charts.setData("not a range")).toBe(false)
    expect(c.charts.setData("=Sheet1!$A$1:$B$4")).toBe(true)
    // Series keep running the way they were switched to: across the rows.
    expect(c.charts.spec()!.series.map((s) => s.name)).toEqual(["Q1", "Q2", "Q3"])
  })

  it("follows a copied sheet, reading the copy", () => {
    const c = sheet(SALES)
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.charts.insert()
    c.sheets.copy(0, 1)
    const copy = c.wb.sheets[1]
    expect(copy.charts[0].refs!.series[0].values).toBe("'Sheet1 (2)'!$B$2:$B$4")
    expect(copy.charts[0].id).not.toBe(c.wb.sheets[0].charts[0].id)
  })
})
