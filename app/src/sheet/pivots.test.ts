import { describe, expect, it, vi } from "vitest"
import { newPivotField, PIVOT_MESSAGE, type PivotTable } from "@cells/core"
import { SheetController } from "./controller"

/** A sheet with a PivotTable in C3:D6 and one report filter above it (C1:D1). */
function book() {
  const c = new SheetController()
  const fields = [newPivotField(2), newPivotField(0)]
  fields[0].axis = "row"
  fields[1].dataField = true
  const table: PivotTable = {
    name: "PivotTable1",
    cacheId: 1,
    at: { row: 2, col: 2 },
    location: { r0: 2, c0: 2, r1: 5, c1: 3 },
    fields,
    rows: [0],
    cols: [],
    pages: [{ field: 0, item: null }],
    values: [{ name: "Sum of Sales", field: 1, fn: "sum" }],
    rowGrandTotals: true,
    colGrandTotals: true,
    compact: true,
    outline: true,
    compactData: true,
    captions: { data: "Values" },
    showError: false,
    showMissing: true,
    style: { name: null, rowHeaders: true, colHeaders: true, rowStripes: false, colStripes: false, lastColumn: true },
  }
  c.sheet.pivots = [table]
  c.sheet.cells.set(3, 2, { value: "North" })
  c.shell.alert = vi.fn()
  return c
}
const select = (c: SheetController, r0: number, c0: number, r1 = r0, c1 = c0) => c.goTo({ r0, c0, r1, c1 })

describe("a PivotTable's cells", () => {
  it("can't be typed over, cleared or pasted into, as in Excel", () => {
    const c = book()
    select(c, 3, 2)
    c.beginEdit("enter", "x")
    expect(c.edit).toBeNull()
    expect(c.shell.alert).toHaveBeenCalledWith(PIVOT_MESSAGE)
    c.clearContents()
    expect(c.sheet.cells.get(3, 2)?.value).toBe("North")
    // The report filter's cells belong to it too.
    select(c, 0, 3)
    c.beginEdit("enter", "x")
    expect(c.edit).toBeNull()
    // Next to it is fine.
    select(c, 3, 4)
    c.beginEdit("enter", "x")
    expect(c.edit).not.toBeNull()
  })

  it("moves with rows inserted above it, and refuses rows inserted through it", () => {
    const c = book()
    select(c, 4, 0)
    c.insertLines("rows")
    expect(c.shell.alert).toHaveBeenCalledWith(PIVOT_MESSAGE)
    expect(c.sheet.pivots[0].location.r0).toBe(2)
    select(c, 0, 0)
    c.insertLines("rows")
    expect(c.sheet.pivots[0].location).toEqual({ r0: 3, c0: 2, r1: 6, c1: 3 })
    expect(c.sheet.pivots[0].at).toEqual({ row: 3, col: 2 })
    c.undo()
    expect(c.sheet.pivots[0].location.r0).toBe(2)
    // Columns to its right don't touch it; deleting all its columns deletes it.
    select(c, 0, 5)
    c.insertLines("cols")
    expect(c.sheet.pivots[0].location.c0).toBe(2)
    select(c, 0, 2, 0, 3)
    c.deleteLines("cols")
    expect(c.sheet.pivots).toEqual([])
    c.undo()
    expect(c.sheet.pivots).toHaveLength(1)
  })
})
