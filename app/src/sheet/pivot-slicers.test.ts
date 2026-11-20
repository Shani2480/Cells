import { describe, expect, it, vi } from "vitest"
import { isChart, pivotLayout, Workbook, type SheetSlicer, type Selection } from "@cells/core"
import { SheetController } from "./controller"

/** Region, Channel, Units on Data!A1:C5; PivotTable1 (Units by Region) at F1 and PivotTable2 (by Channel) at J1. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  c.shell.openDialog = vi.fn()
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Units"],
    ["North", "Shop", 4],
    ["South", "Online", 2],
    ["North", "Online", 5],
    ["East", "Shop", 1],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  const recipe = (rows: number[]) => ({ title: "", rows, cols: [], values: [2] })
  c.pivots.create("A1:C5", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } }, undefined, recipe([0]))
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  c.pivots.create("A1:C5", { kind: "existing", sheet: 0, at: { row: 0, col: 9 } }, undefined, recipe([1]))
  c.goTo({ r0: 1, c0: 5, r1: 1, c1: 5 })
  expect(c.wb.pivotCaches).toHaveLength(1)
  return c
}
const slicers = (c: SheetController) =>
  c.wb.sheets.flatMap((s) => s.objects.filter((o): o is SheetSlicer => !isChart(o) && o.kind === "slicer"))
const total = (c: SheetController, name: string) => {
  const t = c.sheet.pivots.find((x) => x.name === name)!
  const cells = pivotLayout(t, c.wb.pivotCaches[0]).cells
  const label = cells.find((x) => x.value === "Grand Total")!
  return cells.find((x) => x.row === label.row && typeof x.value === "number")!.value
}
const picked = (c: SheetController, o: SheetSlicer) =>
  c.slicers
    .view(o)
    .items.filter((i) => i.selected)
    .map((i) => i.text)

describe("PivotTable slicers", () => {
  it("Insert Slicer in a PivotTable lists its fields and makes slicers connected to it", () => {
    const c = setup()
    c.slicers.openInsert()
    expect(c.shell.openDialog).toHaveBeenCalledWith("insertSlicers")
    expect(c.slicers.insertColumns()).toEqual(["Region", "Channel", "Units"])
    const [channel] = c.slicers.insert(["Channel"])
    expect(channel).toMatchObject({ table: -1, column: "Channel", caption: "Channel" })
    expect(channel.pivot).toEqual({
      cacheId: c.wb.pivotCaches[0].id,
      field: "Channel",
      tables: [{ sheet: "Data", name: "PivotTable1" }],
    })
    expect(c.slicers.view(channel).items.map((i) => i.text)).toEqual(["Online", "Shop"])
  })

  it("a pick filters the PivotTable (the field off its axes), Clear Filter shows all, one undo step each", () => {
    const c = setup()
    const [channel] = c.slicers.insert(["Channel"])
    expect(total(c, "PivotTable1")).toBe(12)
    c.slicers.pick(channel, "x:0", "only")
    expect(picked(c, channel)).toEqual(["Shop"])
    expect(total(c, "PivotTable1")).toBe(5)
    expect(c.sheet.cells.get(1, 5)?.value).toBe("North")
    expect(c.slicers.view(channel).filtered).toBe(true)
    c.undo()
    expect(total(c, "PivotTable1")).toBe(12)
    c.redo()
    c.slicers.clear(slicers(c)[0])
    expect(total(c, "PivotTable1")).toBe(12)
  })

  it("Report Connections adds a PivotTable (it takes the picks) and lets one go (it shows all again)", () => {
    const c = setup()
    const [region] = c.slicers.insert(["Region"])
    c.slicers.pick(region, "x:0", "only") // North
    expect(total(c, "PivotTable1")).toBe(9)
    expect(c.slicers.reportConnections(slicers(c)[0]).map((r) => [r.name, r.on])).toEqual([
      ["PivotTable1", true],
      ["PivotTable2", false],
    ])
    expect(c.slicers.setReportConnections(slicers(c)[0], ["Data!PivotTable1", "Data!PivotTable2"])).toBeNull()
    expect(total(c, "PivotTable2")).toBe(9)
    expect(c.slicers.setReportConnections(slicers(c)[0], [])).toMatch(/at least one/)
    c.slicers.setReportConnections(slicers(c)[0], ["Data!PivotTable2"])
    expect(total(c, "PivotTable1")).toBe(12)
    expect(total(c, "PivotTable2")).toBe(9)
    c.undo()
    expect(total(c, "PivotTable1")).toBe(9)
    expect(slicers(c)[0].pivot?.tables).toHaveLength(2)
  })

  it("Filter Connections lists the slicers of the PivotTable's data and connects them", () => {
    const c = setup()
    c.slicers.insert(["Region"])
    c.goTo({ r0: 1, c0: 9, r1: 1, c1: 9 })
    expect(c.slicers.filterConnections().map((r) => [r.caption, r.on])).toEqual([["Region", false]])
    expect(c.slicers.setFilterConnections([slicers(c)[0].id])).toBe(true)
    expect(slicers(c)[0].pivot?.tables.map((t) => t.name)).toEqual(["PivotTable1", "PivotTable2"])
  })

  it("follows its PivotTable's new name and sheet name, and goes when its only PivotTable is deleted", () => {
    const c = setup()
    c.slicers.insert(["Region"])
    expect(c.pivots.rename("Sales")).toBeNull()
    expect(slicers(c)[0].pivot?.tables).toEqual([{ sheet: "Data", name: "Sales" }])
    c.sheets.rename(0, "Facts")
    expect(slicers(c)[0].pivot?.tables).toEqual([{ sheet: "Facts", name: "Sales" }])
    const area = c.sheet.pivots[0].location!
    expect(c.pivots.deleteCovered([{ ...area, r0: 0 }])).toBe(true)
    expect(slicers(c)).toHaveLength(0)
    c.undo()
    expect(slicers(c)).toHaveLength(1)
  })
})
