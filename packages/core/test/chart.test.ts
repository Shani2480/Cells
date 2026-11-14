// Charts in the workbook: the ranges Insert Chart takes from data (as Excel 365 does), reading them, and charts
// following edits to the sheet.
import { describe, expect, it } from "vitest"
import {
  chartRefsFor,
  copiedSheetCharts,
  readChartData,
  switchChartRefs,
  Workbook,
  writeInputOp,
  type Op,
  type SheetChart,
} from "../src"

function sample() {
  const wb = new Workbook(["Sales", "Other"])
  const rows = [
    ["", "North", "South"],
    ["Q1", "12", "8"],
    ["Q2", "18", "11"],
    ["Q3", "15", "9"],
  ]
  const ops: Op[] = []
  rows.forEach((r, row) => r.forEach((t, col) => t && ops.push(writeInputOp(0, wb.sheets[0], wb.styles, row, col, t))))
  wb.transact(ops, { label: "Setup" })
  return wb
}

const chart = (wb: Workbook): SheetChart => ({
  id: 1,
  name: "Chart 1",
  anchor: { from: { row: 5, col: 1, x: 0, y: 0 }, to: { row: 15, col: 6, x: 0, y: 0 } },
  spec: {},
  refs: chartRefsFor(wb.sheets[0], { r0: 0, c0: 0, r1: 3, c1: 2 }),
})

describe("chart ranges", () => {
  it("take series down the columns when there are more rows, names from the header and categories from the left", () => {
    const wb = sample()
    expect(chartRefsFor(wb.sheets[0], { r0: 0, c0: 0, r1: 3, c1: 2 })).toEqual({
      categories: "Sales!$A$2:$A$4",
      series: [
        { name: "Sales!$B$1", values: "Sales!$B$2:$B$4" },
        { name: "Sales!$C$1", values: "Sales!$C$2:$C$4" },
      ],
    })
    // Numbers only: no names, no categories.
    expect(chartRefsFor(wb.sheets[0], { r0: 1, c0: 1, r1: 3, c1: 1 })).toEqual({
      categories: null,
      series: [{ name: null, values: "Sales!$B$2:$B$4" }],
    })
  })

  it("read as the cells show them; Switch Row/Column takes the series the other way", () => {
    const wb = sample()
    const refs = chartRefsFor(wb.sheets[0], { r0: 0, c0: 0, r1: 3, c1: 2 })
    expect(readChartData(wb.sheets, wb.styles, refs)).toEqual({
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "North", values: [12, 18, 15], texts: ["", "", ""] },
        { name: "South", values: [8, 11, 9], texts: ["", "", ""] },
      ],
    })
    const switched = switchChartRefs(wb.sheets, refs)!
    expect(switched.categories).toBe("Sales!$B$1:$C$1")
    expect(switched.series.map((s) => s.values)).toEqual(["Sales!$B$2:$C$2", "Sales!$B$3:$C$3", "Sales!$B$4:$C$4"])
  })

  it("follow inserted rows, a renamed sheet, and undo", () => {
    const wb = sample()
    wb.transact([{ kind: "charts", sheet: 0, charts: [chart(wb)] }], { label: "Insert Chart" })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 2, count: 1 }], { label: "Insert" })
    const c = wb.sheets[0].charts[0]
    expect(c.refs!.series[0].values).toBe("Sales!$B$2:$B$5")
    expect(c.refs!.categories).toBe("Sales!$A$2:$A$5")
    expect(c.anchor.from.row).toBe(6)
    wb.transact([{ kind: "renameSheet", sheet: 0, name: "Q Data" }], { label: "Rename" })
    expect(wb.sheets[0].charts[0].refs!.series[1].name).toBe("'Q Data'!$C$1")
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].charts[0].refs!.series[0].values).toBe("Sales!$B$2:$B$4")
    expect(wb.sheets[0].charts[0].anchor.from.row).toBe(5)
  })

  it("hidden rows are left out of the data, as Excel plots by default", () => {
    const wb = sample()
    wb.sheets[0].rows.setHidden(2, true)
    const data = readChartData(wb.sheets, wb.styles, chart(wb).refs!)
    expect(data.categories).toEqual(["Q1", "Q3"])
    expect(data.series[0].values).toEqual([12, 15])
  })

  it("several leading text columns are the categories together", () => {
    const wb = new Workbook(["Data"])
    const rows = [
      ["Region", "Product", "Units"],
      ["North", "Pens", "120"],
      ["South", "Ink", "80"],
    ]
    const ops: Op[] = []
    rows.forEach((r, row) => r.forEach((t, col) => ops.push(writeInputOp(0, wb.sheets[0], wb.styles, row, col, t))))
    wb.transact(ops, { label: "Setup" })
    const refs = chartRefsFor(wb.sheets[0], { r0: 0, c0: 0, r1: 2, c1: 2 })
    expect(refs).toEqual({
      categories: "Data!$A$2:$B$3",
      series: [{ name: "Data!$C$1", values: "Data!$C$2:$C$3" }],
    })
    expect(readChartData(wb.sheets, wb.styles, refs).categories).toEqual(["North Pens", "South Ink"])
  })

  it("a copied sheet's charts read the copy, with ids of their own", () => {
    const wb = sample()
    const c = chart(wb)
    c.refs!.categories = "Other!$A$1:$A$3"
    wb.sheets[0].charts = [c]
    const [copy] = copiedSheetCharts(wb.sheets[0].charts, wb.sheets, wb.sheets[0], "Sales (2)")
    expect(copy.id).toBe(2)
    expect(copy.refs!.series[0]).toEqual({ name: "'Sales (2)'!$B$1", values: "'Sales (2)'!$B$2:$B$4" })
    expect(copy.refs!.categories).toBe("Other!$A$1:$A$3")
    expect(wb.sheets[0].charts[0].refs!.series[0].values).toBe("Sales!$B$2:$B$4")
  })
})
