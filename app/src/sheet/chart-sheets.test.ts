import { describe, expect, it } from "vitest"
import { isChart, parseInput, selectRange } from "@cells/core"
import { openXlsx, saveXlsx } from "@cells/io"
import { SheetController } from "./controller"

/** Month and Units in A1:B5 of Sheet1 (mock data), with a column chart of them on the sheet. */
function withChart() {
  const c = new SheetController()
  const rows = [
    ["Month", "Units"],
    ["Jan", "12"],
    ["Feb", "19"],
    ["Mar", "15"],
    ["Apr", "23"],
  ]
  c.wb.transact(
    [
      {
        kind: "setCells",
        sheet: 0,
        cells: rows.flatMap((r, row) => r.map((t, col) => ({ row, col, data: { value: parseInput(t)!.value } }))),
      },
    ],
    { label: "Load" },
  )
  c.setSelection(selectRange({ r0: 0, c0: 0, r1: 4, c1: 1 }))
  const chart = c.charts.insert()!
  return { c, chart }
}

describe("chart sheets and Move Chart", () => {
  it("moves a chart to a chart sheet before its sheet, where it's always selected; undo puts it back", () => {
    const { c, chart } = withChart()
    expect(c.charts.moveToNewSheet("Chart1")).toBe(true)
    expect(c.wb.sheets.map((s) => [s.name, s.chartSheet])).toEqual([
      ["Chart1", true],
      ["Sheet1", false],
    ])
    expect(c.sheetIndex).toBe(0)
    expect(c.sheet.objects).toHaveLength(1)
    expect(c.charts.active()?.id).toBe(chart.id)
    expect(c.wb.sheets[1].charts).toEqual([])
    // The chart sheet's chart can't be deleted on its own.
    c.charts.delete()
    expect(c.sheet.charts).toHaveLength(1)
    c.undo()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1"])
    expect(c.wb.sheets[0].charts.map((x) => x.id)).toEqual([chart.id])
  })

  it("moves a chart sheet's chart back onto a worksheet, and the chart sheet goes", () => {
    const { c } = withChart()
    c.charts.moveToNewSheet("Units")
    expect(c.charts.moveToSheet(1)).toBe(true)
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1"])
    expect(c.sheet.charts).toHaveLength(1)
    expect(c.charts.active()).not.toBeNull()
  })

  it("renames a chart sheet when its chart moves to a new sheet name, and refuses a taken name", () => {
    const { c } = withChart()
    const alerts: string[] = []
    c.shell.alert = (m: string) => void alerts.push(m)
    expect(c.charts.moveToNewSheet("sheet1")).toBe(false)
    expect(alerts[0]).toMatch(/already|taken|name/i)
    c.charts.moveToNewSheet("Units")
    expect(c.charts.moveToNewSheet("Units by month")).toBe(true)
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Units by month", "Sheet1"])
  })

  it("F11 makes a chart of the data on a new chart sheet named Chart1", () => {
    const { c } = withChart()
    c.setSelection(selectRange({ r0: 0, c0: 0, r1: 4, c1: 1 }))
    expect(c.charts.insertSheet()).toBe(true)
    expect(c.wb.sheets.map((s) => [s.name, s.chartSheet])).toEqual([
      ["Chart1", true],
      ["Sheet1", false],
    ])
    const chart = c.sheet.objects.find(isChart)!
    expect(chart.refs?.series[0].values).toBe("Sheet1!$B$2:$B$5")
  })

  it("a chart sheet made in Cells saves and opens again as a chart sheet", async () => {
    const { c } = withChart()
    c.charts.moveToNewSheet("Units")
    const { workbook } = await openXlsx(await saveXlsx(c.wb), "x.xlsx")
    expect(workbook.sheets.map((s) => [s.name, s.chartSheet])).toEqual([
      ["Units", true],
      ["Sheet1", false],
    ])
    expect(workbook.sheets[0].charts[0].refs?.series[0].values).toBe("Sheet1!$B$2:$B$5")
  })
})
