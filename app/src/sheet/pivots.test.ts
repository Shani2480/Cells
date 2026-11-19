import { describe, expect, it, vi } from "vitest"
import {
  newPivotField,
  PIVOT_MESSAGE,
  pivotCacheFrom,
  pivotCellsPlan,
  readPivotSource,
  type PivotTable,
  type Selection,
  Workbook,
} from "@cells/core"
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

describe("Refresh", () => {
  /** Data on Sheet1 (A1:B5), and a PivotTable of Sales by Region on Sheet2 at A3. */
  function setup() {
    const c = new SheetController(new Workbook<Selection>(["Sheet1", "Sheet2"]))
    c.shell.alert = vi.fn()
    const data = c.wb.sheets[0]
    const rows: (string | number)[][] = [
      ["Region", "Sales"],
      ["North", 10],
      ["South", 20],
      ["North", 5],
      ["East", 7],
    ]
    rows.forEach((r, row) => r.forEach((value, col) => data.cells.set(row, col, { value })))
    const source = { kind: "range" as const, sheet: "Sheet1", ref: "A1:B5" }
    const read = readPivotSource(c.wb.sheets, c.wb.styles, source)
    if (typeof read === "string") throw new Error(read)
    const cache = pivotCacheFrom(1, source, read)
    const fields = [newPivotField(3), newPivotField(0)]
    fields[0].axis = "row"
    fields[1].dataField = true
    const table: PivotTable = {
      ...book().sheet.pivots[0],
      pages: [],
      fields,
      at: { row: 2, col: 0 },
      location: { r0: 2, c0: 0, r1: 2, c1: 0 },
    }
    c.wb.pivotCaches = [cache]
    const plan = pivotCellsPlan(c.wb.sheets, c.wb.styles, 1, null, table, cache)
    c.wb.transact([...plan.ops, { kind: "pivots", sheet: 1, tables: [plan.table] }], { label: "Make" })
    c.wb.clearHistory()
    return c
  }
  const column = (c: SheetController, r0: number, r1: number) =>
    Array.from({ length: r1 - r0 + 1 }, (_, i) => [0, 1].map((col) => c.wb.sheets[1].cells.get(r0 + i, col)?.value))

  it("reads the source again and lays the PivotTable out anew, as one step to undo", () => {
    const c = setup()
    expect(column(c, 2, 6)).toEqual([
      ["Row Labels", "Sum of Sales"],
      ["North", 15],
      ["South", 20],
      ["East", 7],
      ["Grand Total", 42],
    ])
    const data = c.wb.sheets[0]
    data.cells.set(2, 0, { value: "West" })
    data.cells.set(5, 0, { value: "North" })
    data.cells.set(5, 1, { value: 100 })
    c.pivots.refreshAll()
    // The source stays A1:B5, as in Excel (a table as the source grows with it).
    expect(c.wb.pivotCaches[0].source).toEqual({ kind: "range", sheet: "Sheet1", ref: "A1:B5" })
    expect(column(c, 2, 6)).toEqual([
      ["Row Labels", "Sum of Sales"],
      ["North", 15],
      ["East", 7],
      ["West", 20],
      ["Grand Total", 42],
    ])
    expect(c.wb.sheets[1].cells.get(7, 0)).toBeUndefined()
    c.undo()
    expect(column(c, 3, 4)).toEqual([
      ["North", 15],
      ["South", 20],
    ])
    expect(c.wb.pivotCaches[0].fields[0].items).toEqual(["North", "South", "East"])
  })

  it("keeps hidden items hidden, clears the rows a smaller table leaves, and drops fields that are gone", () => {
    const c = setup()
    const t = c.wb.sheets[1].pivots[0]
    t.fields[0].items[1].hidden = true
    c.pivots.refreshAll()
    expect(column(c, 2, 5)).toEqual([
      ["Row Labels", "Sum of Sales"],
      ["North", 15],
      ["East", 7],
      ["Grand Total", 22],
    ])
    expect(column(c, 6, 6)).toEqual([[undefined, undefined]])
    c.wb.sheets[0].cells.set(0, 1, { value: "Amount" })
    c.pivots.refreshAll()
    expect(c.wb.sheets[1].pivots[0].values).toEqual([])
    expect(c.wb.sheets[1].cells.get(2, 1)).toBeUndefined()
  })

  it("asks before growing over other data, and says when the source is gone", () => {
    const c = setup()
    c.wb.sheets[1].cells.set(7, 0, { value: "keep?" })
    c.wb.sheets[0].cells.set(5, 0, { value: "West" })
    c.wb.sheets[0].cells.set(5, 1, { value: 1 })
    c.wb.pivotCaches[0].source = { kind: "range", sheet: "Sheet1", ref: "A1:B6" }
    const confirm = vi.fn()
    c.shell.confirm = confirm
    c.pivots.refreshAll()
    expect(confirm).toHaveBeenCalledWith(
      "There's already data in [Book1.xlsx]Sheet2. Do you want to replace it?",
      expect.any(Function),
    )
    expect(c.wb.sheets[1].cells.get(7, 0)?.value).toBe("keep?")
    confirm.mock.calls[0][1]()
    expect(c.wb.sheets[1].cells.get(7, 0)?.value).toBe("Grand Total")
    c.wb.pivotCaches[0].source = { kind: "range", sheet: "Gone", ref: "A1:B6" }
    c.pivots.refreshAll()
    expect(c.shell.alert).toHaveBeenCalledWith("Reference isn't valid.")
  })

  it("Refresh works on the PivotTable at the active cell, past the protection of its cells", () => {
    const c = setup()
    c.activate(1)
    select(c, 3, 0)
    expect(c.pivots.active()?.name).toBe("PivotTable1")
    c.wb.sheets[0].cells.set(1, 1, { value: 11 })
    expect(c.pivots.refresh()).toBe(true)
    expect(c.wb.sheets[1].cells.get(3, 1)?.value).toBe(16)
  })
})

describe("Expand and collapse", () => {
  /** Sales by Region, then Product, in compact form at A3 of Sheet2. */
  function setup() {
    const c = new SheetController(new Workbook<Selection>(["Sheet1", "Sheet2"]))
    c.shell.alert = vi.fn()
    const rows: (string | number)[][] = [
      ["Region", "Product", "Sales"],
      ["North", "Apples", 10],
      ["North", "Pears", 5],
      ["South", "Apples", 20],
    ]
    rows.forEach((r, row) => r.forEach((value, col) => c.wb.sheets[0].cells.set(row, col, { value })))
    const source = { kind: "range" as const, sheet: "Sheet1", ref: "A1:C4" }
    const read = readPivotSource(c.wb.sheets, c.wb.styles, source)
    if (typeof read === "string") throw new Error(read)
    const cache = pivotCacheFrom(1, source, read)
    const fields = [newPivotField(2), newPivotField(2), newPivotField(0)]
    fields[0].axis = "row"
    fields[1].axis = "row"
    fields[2].dataField = true
    const table: PivotTable = {
      ...book().sheet.pivots[0],
      pages: [],
      fields,
      rows: [0, 1],
      values: [{ name: "Sum of Sales", field: 2, fn: "sum" }],
      at: { row: 2, col: 0 },
      location: { r0: 2, c0: 0, r1: 2, c1: 0 },
    }
    c.wb.pivotCaches = [cache]
    const plan = pivotCellsPlan(c.wb.sheets, c.wb.styles, 1, null, table, cache)
    c.wb.transact([...plan.ops, { kind: "pivots", sheet: 1, tables: [plan.table] }], { label: "Make" })
    c.wb.clearHistory()
    c.activate(1)
    return c
  }
  const labels = (c: SheetController) =>
    Array.from({ length: 7 }, (_, r) => c.sheet.cells.get(r + 2, 0)?.value).filter((v) => v !== undefined)

  it("the +/− button and double-clicking an outer item close and open it, each one undo step", () => {
    const c = setup()
    expect(labels(c)).toEqual(["Row Labels", "North", "Apples", "Pears", "South", "Apples", "Grand Total"])
    // North's label (A4) has a − button; the inner items have none.
    expect(c.pivots.cellAt(3, 0)?.cell?.button).toBe("collapse")
    expect(c.pivots.cellAt(4, 0)?.cell?.button).toBeUndefined()
    expect(c.pivots.toggle(4, 0)).toBe(false)
    expect(c.pivots.toggle(3, 0)).toBe(true)
    expect(labels(c)).toEqual(["Row Labels", "North", "South", "Apples", "Grand Total"])
    expect(c.sheet.cells.get(3, 1)?.value).toBe(15)
    expect(c.pivots.cellAt(3, 0)?.cell?.button).toBe("expand")
    // The label that moved up to South's old line has no indent left over from Apples.
    expect(c.wb.styles.get(c.sheet.cells.style(4, 0)).indent ?? 0).toBe(0)
    expect(c.wb.styles.get(c.sheet.cells.style(5, 0)).indent).toBe(1)
    c.undo()
    expect(labels(c)).toEqual(["Row Labels", "North", "Apples", "Pears", "South", "Apples", "Grand Total"])
    expect(c.sheet.pivots[0].fields[0].items.some((i) => i.collapsed)).toBe(false)
  })

  it("the cell menu's Expand/Collapse works on the item or its whole field", () => {
    const c = setup()
    select(c, 3, 0)
    expect(c.pivots.drillState()).toEqual({ expand: false, collapse: true, expandField: false, collapseField: true })
    // On an inner item, only Collapse Entire Field, which closes the field above.
    select(c, 4, 0)
    expect(c.pivots.drillState()).toEqual({ expand: false, collapse: false, expandField: false, collapseField: true })
    expect(c.pivots.drill("item", true)).toBe(false)
    expect(c.pivots.drill("field", true)).toBe(true)
    expect(labels(c)).toEqual(["Row Labels", "North", "South", "Grand Total"])
    select(c, 4, 0)
    expect(c.pivots.drillState()).toMatchObject({ expand: true, expandField: true, collapseField: false })
    expect(c.pivots.drill("field", false)).toBe(true)
    expect(labels(c)).toHaveLength(7)
    // Off a label there's no Expand/Collapse.
    select(c, 3, 1)
    expect(c.pivots.drillState()).toBeNull()
  })

  it("PivotTable Options' Show expand/collapse buttons off hides the buttons", () => {
    const c = setup()
    const t = { ...c.sheet.pivots[0], showDrill: false as const }
    c.sheet.pivots = [t]
    expect(c.pivots.cellAt(3, 0)?.cell?.button).toBeUndefined()
    expect(c.pivots.cellAt(3, 0)?.cell?.item).toMatchObject({ field: 0, axis: "row", level: 0 })
  })
})
