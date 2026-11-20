import { describe, expect, it, vi } from "vitest"
import {
  pivotReportLayout,
  withFieldIn,
  withReportLayout,
  withSubtotalsShown,
  Workbook,
  type Selection,
} from "@cells/core"
import { SheetController } from "./controller"

/** Units by Region and Channel (rows) at F1, from Data!A1:D5; G1 and beyond on Data hold nothing else. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data", "Other"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Units", "Amount"],
    ["North", "Shop", 4, 40],
    ["South", "Online", 2, 20],
    ["North", "Online", 5, 50],
    ["East", "Shop", 1, 10],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.pivots.create("A1:D5", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
  c.pivots.change((t, k) => withFieldIn(withFieldIn(withFieldIn(t, k, 0, "rows"), k, 1, "rows"), k, 2, "values"))
  return c
}

const col = (c: SheetController, col: number, rows = 12) =>
  Array.from({ length: rows }, (_, r) => c.sheet.cells.get(r, col)?.value)

describe("PivotTable Analyze", () => {
  it("renames the PivotTable and the active field", () => {
    const c = setup()
    expect(c.pivots.rename("Sales")).toBeNull()
    expect(c.sheet.pivots[0].name).toBe("Sales")
    expect(c.pivots.rename("")).toBe("PivotTable name is not valid.")
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    expect(c.pivots.activeField()?.name).toBe("Region")
    expect(c.pivots.renameActiveField("Area")).toBeNull()
    expect(c.sheet.pivots[0].fields[0].name).toBe("Area")
    c.goTo({ r0: 1, c0: 6, r1: 1, c1: 6 })
    expect(c.pivots.activeField()?.name).toBe("Sum of Units")
    expect(c.pivots.renameActiveField("Region")).toBe("PivotTable field name already exists.")
  })

  it("Clear Filters shows every item again; Clear All leaves it empty", () => {
    const c = setup()
    c.pivots.showItems(0, 5, 0, new Set([0]))
    expect(col(c, 5)).not.toContain("South")
    c.pivots.clear("filters")
    expect(col(c, 5)).toContain("South")
    c.pivots.clear("all")
    expect(c.sheet.pivots[0].rows).toEqual([])
    expect(col(c, 5).filter((v) => v !== undefined)).toEqual([])
    c.undo()
    expect(col(c, 5)).toContain("South")
  })

  it("selects its values, labels and the whole table", () => {
    const c = setup()
    // F1 Row Labels | Sum of Units; F2.. the lines; Grand Total at row 9.
    c.pivots.select("values")
    expect(c.selection.ranges).toEqual([{ r0: 1, c0: 6, r1: 8, c1: 6 }])
    c.pivots.select("labels")
    expect(c.selection.ranges[0]).toEqual({ r0: 0, c0: 5, r1: 8, c1: 5 })
    c.pivots.select("entire")
    expect(c.selection.ranges).toEqual([{ r0: 0, c0: 5, r1: 8, c1: 6 }])
  })

  it("moves the PivotTable on its sheet and to another, in one undo step", () => {
    const c = setup()
    expect(c.pivots.move({ kind: "existing", sheet: 0, at: { row: 20, col: 1 } })).toBeNull()
    expect(c.sheet.pivots[0].at).toEqual({ row: 20, col: 1 })
    expect(c.sheet.cells.get(20, 1)?.value).toBe("Row Labels")
    expect(c.sheet.cells.get(0, 5)).toBeUndefined()
    c.pivots.move({ kind: "existing", sheet: 1, at: { row: 2, col: 0 } })
    expect(c.sheet.name).toBe("Other")
    expect(c.sheet.cells.get(2, 0)?.value).toBe("Row Labels")
    expect(c.wb.sheets[0].pivots).toEqual([])
    expect(c.wb.sheets[0].cells.get(20, 1)).toBeUndefined()
    c.undo()
    expect(c.wb.sheets[0].pivots).toHaveLength(1)
    expect(c.wb.sheets[1].pivots).toHaveLength(0)
    expect(c.wb.sheets[0].cells.get(20, 1)?.value).toBe("Row Labels")
  })

  it("Change Data Source reads the new range, keeping the fields it still has", () => {
    const c = setup()
    expect(c.pivots.sourceText()).toBe("Data!$A$1:$D$5")
    expect(c.pivots.changeSource("Data!A1:D3")).toBeNull()
    expect(c.wb.pivotCaches[0].source).toEqual({ kind: "range", sheet: "Data", ref: "A1:D3" })
    expect(col(c, 5).filter((v) => v !== undefined)).toEqual([
      "Row Labels",
      "North",
      "Shop",
      "South",
      "Online",
      "Grand Total",
    ])
    expect(c.pivots.changeSource("Nowhere!A1")).toBe("Reference isn't valid.")
  })
})

describe("Design", () => {
  it("Report Layout, Subtotals and Grand Totals change every field", () => {
    const c = setup()
    c.pivots.layout((t) => withReportLayout(t, "tabular"), "Report Layout")
    expect(pivotReportLayout(c.sheet.pivots[0])).toBe("tabular")
    // Tabular: Region and Channel in their own columns.
    expect([5, 6, 7].map((k) => c.sheet.cells.get(0, k)?.value)).toEqual(["Region", "Channel", "Sum of Units"])
    c.pivots.layout((t) => withSubtotalsShown(t, "none"), "Subtotals")
    expect(col(c, 5)).not.toContain("North Total")
    c.pivots.setStyleOption("rowStripes", true)
    expect(c.sheet.pivots[0].style.rowStripes).toBe(true)
    c.pivots.setStyle("PivotStyleMedium9")
    expect(c.sheet.pivots[0].style.name).toBe("PivotStyleMedium9")
  })
})

describe("Delete on a PivotTable", () => {
  it("deletes a wholly selected PivotTable; part of one gets Excel's message", () => {
    const c = setup()
    c.goTo({ r0: 1, c0: 5, r1: 3, c1: 6 })
    c.clearContents()
    expect(c.shell.alert).toHaveBeenCalledWith(expect.stringMatching(/affect a PivotTable/))
    expect(c.sheet.pivots).toHaveLength(1)
    c.goTo({ r0: 0, c0: 4, r1: 20, c1: 8 })
    c.clearContents()
    expect(c.sheet.pivots).toEqual([])
    expect(c.wb.pivotCaches).toEqual([])
    expect(col(c, 5).filter((v) => v !== undefined)).toEqual([])
    c.undo()
    expect(c.sheet.pivots).toHaveLength(1)
    expect(c.wb.pivotCaches).toHaveLength(1)
    expect(col(c, 5)[0]).toBe("Row Labels")
  })
})

describe("GETPIVOTDATA", () => {
  it("follows its PivotTable: a refresh and a change of layout recalculate it", () => {
    const c = setup()
    const formula = (input: string) =>
      c.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 15, col: 0, data: { input, value: 0 } }] }], {
        label: "Formula",
      })
    formula('=GETPIVOTDATA("Sum of Units",F1,"Region","North")')
    expect(c.sheet.cells.get(15, 0)?.value).toBe(9)
    // A source change shows once the PivotTable is refreshed.
    c.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 1, col: 2, data: { value: 14 } }] }], { label: "Edit" })
    expect(c.sheet.cells.get(15, 0)?.value).toBe(9)
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    c.pivots.refresh()
    expect(c.sheet.cells.get(15, 0)?.value).toBe(19)
    // Region subtotals off: the North line holds no value any more.
    c.pivots.layout((t) => withSubtotalsShown(t, "none"), "Subtotals")
    expect(c.sheet.cells.get(15, 0)?.value).toEqual({ error: "#REF!" })
    c.undo()
    expect(c.sheet.cells.get(15, 0)?.value).toBe(19)
  })
})

describe("Generate GetPivotData", () => {
  it("pointing at a value cell writes GETPIVOTDATA; at a label, or with it off, a reference", () => {
    const c = setup()
    localStorage.removeItem("cells.generateGetPivotData")
    c.goTo({ r0: 15, c0: 0, r1: 15, c1: 0 })
    c.beginEdit("enter", "=")
    // G2: North's subtotal; G3: North / Shop; G9: the grand total.
    c.pointAt({ row: 1, col: 6 })
    expect(c.edit?.text).toBe('=GETPIVOTDATA("Sum of Units",$F$1,"Region","North")')
    c.pointAt({ row: 2, col: 6 })
    expect(c.edit?.text).toBe('=GETPIVOTDATA("Sum of Units",$F$1,"Region","North","Channel","Shop")')
    c.pointAt({ row: 8, col: 6 })
    expect(c.edit?.text).toBe('=GETPIVOTDATA("Sum of Units",$F$1)')
    c.pointAt({ row: 2, col: 5 })
    expect(c.edit?.text).toBe("=F3")
    c.pointAt({ row: 2, col: 6 })
    c.commitEdit()
    expect(c.sheet.cells.get(15, 0)?.value).toBe(4)
    c.pivots.generateGetPivotData = false
    c.goTo({ r0: 16, c0: 0, r1: 16, c1: 0 })
    c.beginEdit("enter", "=")
    c.pointAt({ row: 2, col: 6 })
    expect(c.edit?.text).toBe("=G3")
    c.pivots.generateGetPivotData = true
  })
})

describe("Calculated Field", () => {
  it("adds a field to the active PivotTable's values, modifies and deletes it, each one undo step", () => {
    const c = setup()
    c.goTo({ r0: 1, c0: 6, r1: 1, c1: 6 })
    expect(c.pivots.setCalcField("Price", "=Amount/Units")).toBeNull()
    expect(c.pivots.calcFields()).toEqual([{ name: "Price", formula: "Amount/Units" }])
    // F1 Row Labels | Sum of Units | Sum of Price; North (90/9) first, East (10/1) on row 7.
    expect(c.sheet.cells.get(0, 7)?.value).toBe("Sum of Price")
    expect(c.sheet.cells.get(6, 7)?.value).toBe(10)
    expect(c.pivots.setCalcField("price", "=Amount*2")).toBeNull()
    expect(c.sheet.cells.get(6, 7)?.value).toBe(20)
    expect(c.pivots.setCalcField("Units", "=1")).toBe("PivotTable field name already exists.")
    expect(c.pivots.setCalcField("Cost", "=Amount+A1")).toBe(
      "References, names, and arrays are not supported in PivotTable formulas.",
    )
    c.pivots.deleteCalcField("Price")
    expect(c.pivots.calcFields()).toEqual([])
    expect(c.sheet.cells.get(0, 7)).toBeUndefined()
    c.undo()
    expect(c.sheet.cells.get(6, 7)?.value).toBe(20)
    c.undo()
    expect(c.sheet.cells.get(6, 7)?.value).toBe(10)
    c.undo()
    expect(c.wb.pivotCaches[0].fields).toHaveLength(4)
    expect(c.sheet.cells.get(0, 7)).toBeUndefined()
  })

  it("follows a refresh of its data", () => {
    const c = setup()
    c.goTo({ r0: 1, c0: 6, r1: 1, c1: 6 })
    c.pivots.setCalcField("Price", "=Amount/Units")
    c.wb.sheets[0].cells.set(4, 3, { value: 30 })
    c.pivots.refresh()
    expect(c.sheet.cells.get(6, 7)?.value).toBe(30)
  })
})

describe("Grouping", () => {
  /** Units by Day (rows) at F1, from Data!A1:C7: dates over two years, a product, units. */
  function dated() {
    const c = new SheetController(new Workbook<Selection>(["Data"]))
    c.shell.alert = vi.fn()
    c.shell.openDialog = vi.fn()
    const rows: (string | number)[][] = [
      ["Day", "Product", "Units"],
      [45658, "Apples", 1], // 2025-01-01
      [45700, "Pears", 2], // 2025-02-12
      [45720, "Plums", 3], // 2025-03-04
      [46037, "Apples", 4], // 2026-01-15
      [46083, "Pears", 5], // 2026-03-02
      [46174, "Plums", 6], // 2026-06-01
    ]
    rows.forEach((r, row) =>
      r.forEach((value, col) =>
        c.sheet.cells.set(row, col, row > 0 && col === 0 ? { value, style: dateStyle(c) } : { value }),
      ),
    )
    c.pivots.create("A1:C7", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
    c.pivots.change((t, k) => withFieldIn(withFieldIn(t, k, 0, "rows"), k, 2, "values"))
    return c
  }
  const dateStyle = (c: SheetController) => c.wb.styles.intern({ ...c.wb.styles.get(0), numberFormat: "m/d/yyyy" })

  it("groups dates by years and months: the Day field gives its place to Years (Day) and Months (Day)", () => {
    const c = dated()
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    expect(c.pivots.groupDialogField()).toBe(0)
    expect(c.pivots.groupDialog()).toMatchObject({ dates: true, start: null, end: null, parts: ["months"] })
    c.pivots.groupSelection()
    expect(c.shell.openDialog).toHaveBeenCalledWith("pivotGroup")
    expect(c.pivots.setGroup({ start: null, end: null, by: 1, parts: ["months", "years"] })).toBeNull()
    const names = c.wb.pivotCaches[0].fields.map((f) => f.name)
    expect(names).toEqual(["Day", "Product", "Units", "Months (Day)", "Years (Day)"])
    expect(c.sheet.pivots[0].rows).toEqual([4, 3])
    expect(col(c, 5, 10)).toEqual([
      "Row Labels",
      "2025",
      "Jan",
      "Feb",
      "Mar",
      "2026",
      "Jan",
      "Mar",
      "Jun",
      "Grand Total",
    ])
    expect(col(c, 6, 10)).toEqual(["Sum of Units", 6, 1, 2, 3, 15, 4, 5, 6, 21])
    // Ungroup from a month: Day comes back in their place.
    c.goTo({ r0: 2, c0: 5, r1: 2, c1: 5 })
    expect(c.pivots.canUngroup()).toBe(true)
    c.pivots.ungroup()
    expect(c.wb.pivotCaches[0].fields).toHaveLength(3)
    expect(c.sheet.pivots[0].rows).toEqual([0])
    c.undo()
    expect(c.sheet.pivots[0].rows).toEqual([4, 3])
    c.undo()
    expect(c.sheet.pivots[0].rows).toEqual([0])
    expect(c.wb.pivotCaches[0].fields).toHaveLength(3)
  })

  it("groups numbers in ranges on the field itself, and ungroups them", () => {
    const c = dated()
    c.pivots.change((t, k) => withFieldIn(withFieldIn(t, k, 0, "filters"), k, 2, "rows"))
    // Units in rows (and values): F4 is its first label under the report filter.
    const at = c.sheet.pivots[0]
    c.goTo({ r0: at.location.r0 + 1, c0: 5, r1: at.location.r0 + 1, c1: 5 })
    expect(c.pivots.groupDialog()).toMatchObject({
      dates: false,
      start: null,
      end: null,
      by: 10,
      autoStart: 1,
      autoEnd: 6,
    })
    expect(c.pivots.setGroup({ start: 1, end: 6, by: 2 })).toBeNull()
    const rows = (k: SheetController) => col(k, 5, 14).filter((v) => typeof v === "string" && /^\d/.test(v))
    expect(rows(c)).toEqual(["1-2", "3-4", "5-6"])
    expect(c.pivots.setGroup({ start: 0, end: 10, by: 5 })).toBeNull()
    // As Excel: the last range reaches the end (its 13-17).
    expect(rows(c)).toEqual(["0-4", "5-10"])
    c.pivots.ungroup()
    expect(c.wb.pivotCaches[0].fields[2].group).toBeUndefined()
    expect(rows(c)).toEqual([])
    expect(col(c, 5, 14).filter((v) => typeof v === "number")).toEqual([1, 2, 3, 4, 5, 6])
  })

  it("Group Selection gathers text items into Product2, and Ungroup splits them again", () => {
    const c = dated()
    c.pivots.change((t, k) => withFieldIn(withFieldIn(t, k, 0, "rows", 99), k, 1, "rows", 0))
    // F2.. Apples, its days, Pears …: select the Apples and Pears labels.
    const labels = col(c, 5, 20)
    const apples = labels.indexOf("Apples")
    const pears = labels.indexOf("Pears")
    c.selection = {
      ...c.selection,
      ranges: [
        { r0: apples, c0: 5, r1: apples, c1: 5 },
        { r0: pears, c0: 5, r1: pears, c1: 5 },
      ],
    }
    expect(c.pivots.groupTarget()?.xs.sort()).toEqual([0, 1])
    expect(c.pivots.groupSelection()).toBe(true)
    const k = c.wb.pivotCaches[0]
    expect(k.fields[3]).toMatchObject({ name: "Product2", group: { kind: "discrete", items: ["Plums", "Group1"] } })
    expect(c.sheet.pivots[0].rows.slice(0, 2)).toEqual([3, 1])
    expect(col(c, 5, 3)).toEqual(["Row Labels", "Group1", "Apples"])
    // Ungroup on Group1: Product2 goes.
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    c.pivots.ungroup()
    expect(c.wb.pivotCaches[0].fields).toHaveLength(3)
    expect(c.sheet.pivots[0].rows.slice(0, 1)).toEqual([1])
  })
})

describe("PivotCharts", () => {
  it("Insert › PivotChart charts the active PivotTable: its lines as categories, outer rows as outer levels", () => {
    const c = setup()
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    const chart = c.pivots.insertChart("column")!
    expect(chart.pivot).toEqual({ sheet: "Data", name: c.sheet.pivots[0].name })
    expect(chart.refs).toBeNull()
    const spec = c.charts.liveSpec(chart)
    expect(spec.title).toEqual({ show: true, text: "Total" })
    // In the PivotTable's order: North (Shop, Online), South, East.
    expect(spec.categories).toEqual(["Shop", "Online", "Online", "Shop"])
    expect(spec.levels).toEqual([["North", "North", "South", "East"]])
    expect(spec.series.map((s) => [s.name, s.values])).toEqual([["Total", [4, 5, 2, 1]]])
    c.undo()
    expect(c.sheet.charts).toEqual([])
  })

  it("follows the PivotTable: its fields, its new name, and its sheet's new name", () => {
    const c = setup()
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    const id = c.pivots.insertChart()!.id
    const chart = () => c.sheet.charts.find((x) => x.id === id)!
    // Channel moves to the columns: a series per channel.
    c.pivots.change((t, k) => withFieldIn(t, k, 1, "cols"))
    const spec = c.charts.liveSpec(chart())
    expect(spec.categories).toEqual(["North", "South", "East"])
    expect(spec.levels).toBeUndefined()
    expect(spec.series.map((s) => [s.name, s.values])).toEqual([
      ["Shop", [4, 0, 1]],
      ["Online", [5, 2, 0]],
    ])
    expect(spec.series[0].blanks).toEqual([1])
    expect(c.pivots.rename("Sales")).toBeNull()
    expect(chart().pivot).toEqual({ sheet: "Data", name: "Sales" })
    c.undo()
    expect(chart().pivot?.name).not.toBe("Sales")
    c.sheets.rename(0, "Facts")
    expect(chart().pivot?.sheet).toBe("Facts")
    expect(c.charts.liveSpec(chart()).categories).toEqual(["North", "South", "East"])
  })

  it("deleting the PivotTable leaves its chart with the points it showed, as Excel does", () => {
    const c = setup()
    c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
    const id = c.pivots.insertChart()!.id
    expect(c.pivots.deleteCovered([{ r0: 0, c0: 5, r1: 20, c1: 8 }])).toBe(true)
    const chart = c.sheet.charts.find((x) => x.id === id)!
    expect(chart.pivot).toBeUndefined()
    const spec = c.charts.liveSpec(chart)
    expect(spec.categories).toEqual(["Shop", "Online", "Online", "Shop"])
    expect(spec.levels).toEqual([["North", "North", "South", "East"]])
    expect(spec.series[0].values).toEqual([4, 5, 2, 1])
    c.undo()
    expect(c.sheet.charts.find((x) => x.id === id)!.pivot?.sheet).toBe("Data")
  })

  it("outside a PivotTable, Create PivotChart makes a PivotTable and a chart of it", () => {
    const c = setup()
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    expect(c.pivots.createWithChart("A1:D5", { kind: "existing", sheet: 1, at: { row: 2, col: 0 } }, "pie")).toBeNull()
    expect(c.sheetIndex).toBe(1)
    const [chart] = c.sheet.charts
    expect(chart.pivot).toEqual({ sheet: "Other", name: c.sheet.pivots[0].name })
    // An empty PivotTable: nothing to draw yet.
    expect(c.charts.liveSpec(chart).type).toBe("pie")
    c.pivots.change((t, k) => withFieldIn(withFieldIn(t, k, 1, "rows"), k, 3, "values"))
    const spec = c.charts.liveSpec(c.sheet.charts[0])
    expect(spec.categories).toEqual(["Shop", "Online"])
    expect(spec.series[0].values).toEqual([50, 70])
  })
})
