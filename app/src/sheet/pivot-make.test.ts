import { describe, expect, it, vi } from "vitest"
import { withFieldIn, withFieldTicked, Workbook, type Selection } from "@cells/core"
import { SheetController } from "./controller"

/** Region, Channel, Units on Data!A1:C5, the active cell inside them. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Units"],
    ["North", "Shop", 4],
    ["South", "Online", 2],
    ["North", "Online", 5],
    ["East", "Shop", 1],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  return c
}

describe("Insert › PivotTable", () => {
  it("offers the data around the active cell, or the table it is in", () => {
    const c = setup()
    expect(c.pivots.defaultSource()).toBe("Data!$A$1:$C$5")
    c.tables.create({ r0: 0, c0: 0, r1: 4, c1: 2 }, true, "TableStyleMedium2")
    expect(c.pivots.defaultSource()).toBe("Table1")
    c.goTo({ r0: 20, c0: 20, r1: 20, c1: 20 })
    expect(c.pivots.defaultSource()).toBe("")
  })

  it("makes an empty PivotTable at A3 of a new sheet before this one, and opens the Fields pane on it", () => {
    const c = setup()
    expect(c.pivots.create("Data!$A$1:$C$5", { kind: "newSheet" })).toBeNull()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet2", "Data"])
    expect(c.sheet.name).toBe("Sheet2")
    expect(c.sheet.pivots[0]).toMatchObject({ name: "PivotTable1", at: { row: 2, col: 0 } })
    expect(c.wb.pivotCaches).toHaveLength(1)
    expect(c.wb.pivotCaches[0].source).toEqual({ kind: "range", sheet: "Data", ref: "A1:C5" })
    expect(c.selection.active).toEqual({ row: 2, col: 0 })
    expect(c.pivots.fieldsShown(false)).toBe(true)
    expect(c.pivots.fieldsShown(true)).toBe(true)
    c.undo()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Data"])
    expect(c.wb.pivotCaches).toHaveLength(0)
  })

  it("ticking fields lays the PivotTable out, one undo step each", () => {
    const c = setup()
    c.pivots.create("A1:C5", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
    expect(c.selection.active).toEqual({ row: 0, col: 5 })
    c.pivots.change((t, k) => withFieldTicked(t, k, 0))
    c.pivots.change((t, k) => withFieldTicked(t, k, 2))
    const cols = (r: number) => [5, 6].map((col) => c.sheet.cells.get(r, col)?.value)
    expect([0, 1, 2, 3, 4].map(cols)).toEqual([
      ["Row Labels", "Sum of Units"],
      ["North", 9],
      ["South", 2],
      ["East", 1],
      ["Grand Total", 12],
    ])
    c.pivots.change((t, k) => withFieldIn(t, k, 1, "cols"))
    expect(c.sheet.cells.get(0, 5)?.value).toBe("Sum of Units")
    c.undo()
    expect(c.sheet.cells.get(0, 5)?.value).toBe("Row Labels")
    c.undo()
    c.undo()
    // Back to empty: nothing in the cells, the PivotTable kept.
    expect(c.sheet.cells.get(0, 5)).toBeUndefined()
    expect(c.sheet.pivots).toHaveLength(1)
  })

  it("says so when the source has no headings, or a PivotTable is in the way", () => {
    const c = setup()
    expect(c.pivots.create("Data!E1:F3", { kind: "newSheet" })).toMatch(/PivotTable field name is not valid/)
    expect(c.pivots.create("nowhere!A1", { kind: "newSheet" })).toBe("Reference isn't valid.")
    c.pivots.create("A1:C5", { kind: "existing", sheet: 0, at: { row: 10, col: 0 } })
    expect(c.pivots.create("A1:C5", { kind: "existing", sheet: 0, at: { row: 12, col: 1 } })).toBe(
      "A PivotTable report cannot overlap another PivotTable report.",
    )
    expect(c.sheet.pivots).toHaveLength(1)
  })

  it("on a phone the pane shows only when asked for", () => {
    const c = setup()
    c.pivots.create("A1:C5", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
    c.pivots.setFieldsOpen(false)
    expect(c.pivots.fieldsShown(true)).toBe(false)
    c.pivots.fieldsOpen = true
    expect(c.pivots.fieldsShown(false)).toBe(true)
    expect(c.pivots.fieldsShown(true)).toBe(false)
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    expect(c.pivots.fieldsShown(false)).toBe(false)
  })
})
