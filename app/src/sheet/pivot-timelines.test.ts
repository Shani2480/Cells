import { describe, expect, it, vi } from "vitest"
import { dateSerial, isChart, pivotLayout, Workbook, type Selection, type SheetTimeline } from "@cells/core"
import { objectSvg } from "@/drawing/render"
import { SheetController } from "./controller"

/** Date, Units on Data!A1:B6 (Jan to May 2025); PivotTable1 (Units, nothing on rows) at E1 and PivotTable2 at H1. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  c.shell.openDialog = vi.fn()
  const date = c.wb.styles.intern({ ...c.wb.styles.get(0), numberFormat: "m/d/yyyy" })
  const rows: [number, number][] = [
    [dateSerial(2025, 1, 10), 1],
    [dateSerial(2025, 2, 3), 2],
    [dateSerial(2025, 3, 31), 4],
    [dateSerial(2025, 4, 1), 8],
    [dateSerial(2025, 5, 20), 16],
  ]
  c.sheet.cells.set(0, 0, { value: "Date" })
  c.sheet.cells.set(0, 1, { value: "Units" })
  rows.forEach(([d, u], i) => {
    c.sheet.cells.set(i + 1, 0, { value: d, style: date })
    c.sheet.cells.set(i + 1, 1, { value: u })
  })
  const recipe = { title: "", rows: [], cols: [], values: [1] }
  c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
  c.pivots.create("A1:B6", { kind: "existing", sheet: 0, at: { row: 0, col: 4 } }, undefined, recipe)
  c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
  c.pivots.create("A1:B6", { kind: "existing", sheet: 0, at: { row: 0, col: 7 } }, undefined, recipe)
  c.goTo({ r0: 1, c0: 4, r1: 1, c1: 4 })
  return c
}
const timelines = (c: SheetController) =>
  c.wb.sheets.flatMap((s) => s.objects.filter((o): o is SheetTimeline => !isChart(o) && o.kind === "timeline"))
const total = (c: SheetController, name: string) => {
  const t = c.sheet.pivots.find((x) => x.name === name)!
  const cells = pivotLayout(t, c.wb.pivotCaches[0]).cells
  return cells.filter((x) => typeof x.value === "number").at(-1)!.value
}
const months = (c: SheetController, o: SheetTimeline) =>
  c.timelines
    .view(o)
    .periods.map((p, i) => [i, p.label] as const)
    .slice(0, 6)

describe("timelines", () => {
  it("Insert Timeline lists the PivotTable's date fields and makes a timeline of months", () => {
    const c = setup()
    expect(c.wb.pivotCaches).toHaveLength(1)
    c.timelines.openInsert()
    expect(c.shell.openDialog).toHaveBeenCalledWith("insertTimelines")
    expect(c.timelines.insertFields()).toEqual(["Date"])
    const [t] = c.timelines.insert(["Date"])
    expect(t).toMatchObject({ kind: "timeline", caption: "Date", level: "months", style: "TimeSlicerStyleLight1" })
    expect(t.pivot.tables).toEqual([{ sheet: "Data", name: "PivotTable1" }])
    expect(months(c, t)).toEqual([
      [0, "JAN"],
      [1, "FEB"],
      [2, "MAR"],
      [3, "APR"],
      [4, "MAY"],
      [5, "JUN"],
    ])
    expect(c.timelines.view(t).periods).toHaveLength(12)
    expect(c.timelines.view(t).label).toBe("All Periods")
    const svg = objectSvg(t, 453, 160, c.wb.theme)
    for (const text of ["Date", "All Periods", "MONTHS", "2025", "JAN"]) expect(svg).toContain(`>${text}<`)
  })

  it("a click picks a month, Shift+click a run, Clear Filter all, as date filters one undo step each", () => {
    const c = setup()
    const [t] = c.timelines.insert(["Date"])
    expect(total(c, "PivotTable1")).toBe(31)
    c.timelines.pick(timelines(c)[0], 2, false)
    expect(total(c, "PivotTable1")).toBe(4)
    expect(c.sheet.pivots[0].filters).toEqual([
      { field: 0, id: 1, kind: "date", op: "between", value: dateSerial(2025, 3, 1), value2: dateSerial(2025, 3, 31) },
    ])
    expect(c.timelines.view(timelines(c)[0]).label).toBe("Mar 2025")
    c.timelines.pick(timelines(c)[0], 4, true)
    expect(total(c, "PivotTable1")).toBe(28)
    expect(c.timelines.view(timelines(c)[0]).label).toBe("Mar - May 2025")
    c.undo()
    expect(total(c, "PivotTable1")).toBe(4)
    c.timelines.clear(timelines(c)[0])
    expect(total(c, "PivotTable1")).toBe(31)
    expect(c.sheet.pivots[0].filters).toBeUndefined()
    // Years: one period, labelled at the level it was picked.
    c.timelines.setLevel(timelines(c)[0], "quarters")
    expect(c.timelines.view(timelines(c)[0]).periods.map((p) => p.label)).toEqual(["Q1", "Q2", "Q3", "Q4"])
    c.timelines.pick(timelines(c)[0], 1, false)
    expect(total(c, "PivotTable1")).toBe(24)
    expect(c.timelines.view(timelines(c)[0]).label).toBe("Q2 2025")
    void t
  })

  it("Report Connections gives another PivotTable the dates, and takes them back off", () => {
    const c = setup()
    c.timelines.insert(["Date"])
    c.timelines.pick(timelines(c)[0], 0, false)
    expect(total(c, "PivotTable2")).toBe(31)
    expect(c.timelines.reportConnections(timelines(c)[0]).map((r) => [r.name, r.on])).toEqual([
      ["PivotTable1", true],
      ["PivotTable2", false],
    ])
    c.timelines.setReportConnections(timelines(c)[0], ["Data!PivotTable1", "Data!PivotTable2"])
    expect(total(c, "PivotTable2")).toBe(1)
    c.timelines.setReportConnections(timelines(c)[0], ["Data!PivotTable2"])
    expect(total(c, "PivotTable1")).toBe(31)
    expect(total(c, "PivotTable2")).toBe(1)
    // Filter Connections lists it beside slicers.
    c.goTo({ r0: 1, c0: 7, r1: 1, c1: 7 })
    expect(c.slicers.filterConnections().map((r) => [r.caption, r.on])).toEqual([["Date", true]])
  })

  it("the Timeline tab's Show boxes, caption and style; follows a renamed PivotTable; goes with it", () => {
    const c = setup()
    c.timelines.insert(["Date"])
    c.timelines.setShown("scrollbar", false)
    c.timelines.setCaption("Order date")
    c.timelines.setStyle("TimeSlicerStyleDark2")
    expect(timelines(c)[0]).toMatchObject({ scrollbar: false, caption: "Order date", style: "TimeSlicerStyleDark2" })
    c.goTo({ r0: 1, c0: 4, r1: 1, c1: 4 })
    expect(c.pivots.rename("Sales")).toBeNull()
    expect(timelines(c)[0].pivot.tables).toEqual([{ sheet: "Data", name: "Sales" }])
    const area = c.sheet.pivots[0].location!
    expect(c.pivots.deleteCovered([{ ...area, r0: 0 }])).toBe(true)
    expect(timelines(c)).toHaveLength(0)
  })
})
