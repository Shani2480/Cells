import { describe, expect, it, vi } from "vitest"
import {
  dateSerial,
  newPivotField,
  pivotCacheFrom,
  pivotCellsPlan,
  readPivotSource,
  Workbook,
  type PivotTable,
  type Selection,
} from "@cells/core"
import { SheetController } from "./controller"

/** Sales by Region (rows) and Channel (columns) on Sheet2 at A3, from Data!A1:D6 with a date column. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data", "Sheet2"]))
  c.shell.alert = vi.fn()
  const data = c.wb.sheets[0]
  const date = c.wb.styles.intern({ ...c.wb.styles.get(0), numberFormat: "m/d/yyyy" })
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Day", "Sales"],
    ["North", "Shop", dateSerial(2025, 1, 2), 10],
    ["South", "Online", dateSerial(2025, 1, 3), 20],
    ["North", "Online", dateSerial(2025, 1, 4), 5],
    ["North", "Shop", dateSerial(2025, 1, 5), 7],
    ["East", "Shop", dateSerial(2025, 1, 6), 1],
  ]
  rows.forEach((r, row) =>
    r.forEach((value, col) => data.cells.set(row, col, { value, ...(col === 2 && row ? { style: date } : {}) })),
  )
  const source = { kind: "range" as const, sheet: "Data", ref: "A1:D6" }
  const read = readPivotSource(c.wb.sheets, c.wb.styles, source)
  if (typeof read === "string") throw new Error(read)
  const cache = pivotCacheFrom(1, source, read)
  const fields = cache.fields.map((f) => newPivotField(f.items?.length ?? 0))
  fields[0].axis = "row"
  fields[1].axis = "col"
  fields[3].dataField = true
  const table: PivotTable = {
    name: "PivotTable1",
    cacheId: 1,
    at: { row: 2, col: 0 },
    location: { r0: 2, c0: 0, r1: 2, c1: 0 },
    fields,
    rows: [0],
    cols: [1],
    pages: [],
    values: [{ name: "Sum of Sales", field: 3, fn: "sum" }],
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
  c.wb.pivotCaches = [cache]
  const plan = pivotCellsPlan(c.wb.sheets, c.wb.styles, 1, null, table, cache)
  c.wb.transact([...plan.ops, { kind: "pivots", sheet: 1, tables: [plan.table] }], { label: "Make" })
  c.wb.clearHistory()
  c.activate(1)
  return c
}

const grid = (c: SheetController, rows: number, cols = 4) =>
  Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, col) => c.sheet.cells.get(r, col)?.value))

describe("Show Details", () => {
  it("double-clicking a value lists its source rows as a table on a new sheet before the PivotTable's", () => {
    const c = setup()
    // A3: Sum of Sales | Column Labels; A4: Row Labels | Shop | Online | Grand Total; A5: North.
    expect(grid(c, 8).slice(3)).toEqual([
      ["Row Labels", "Shop", "Online", "Grand Total"],
      ["North", 17, 5, 22],
      ["South", undefined, 20, 20],
      ["East", 1, undefined, 1],
      ["Grand Total", 18, 25, 43],
    ])
    expect(c.pivots.canShowDetails(4, 1)).toBe(true)
    expect(c.pivots.canShowDetails(4, 0)).toBe(false)
    expect(c.pivots.showDetails(4, 0)).toBe(false)
    expect(c.pivots.showDetails(4, 1)).toBe(true)
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Data", "Sheet3", "Sheet2"])
    expect(c.sheet.name).toBe("Sheet3")
    expect(grid(c, 4)).toEqual([
      ["Region", "Channel", "Day", "Sales"],
      ["North", "Shop", dateSerial(2025, 1, 2), 10],
      ["North", "Shop", dateSerial(2025, 1, 5), 7],
      [undefined, undefined, undefined, undefined],
    ])
    expect(c.wb.styles.get(c.sheet.cells.style(1, 2)).numberFormat).toBe("m/d/yyyy")
    expect(c.sheet.tables[0]).toMatchObject({
      name: "Table1",
      style: "TableStyleMedium2",
      range: { r0: 0, c0: 0, r1: 2, c1: 3 },
    })
    expect(c.selection.ranges[0]).toEqual({ r0: 0, c0: 0, r1: 2, c1: 3 })
    c.undo()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Data", "Sheet2"])
  })

  it("the Grand Total shows every row counted, and filters narrow it", () => {
    const c = setup()
    c.pivots.showDetails(7, 3)
    expect(c.sheet.tables[0].range.r1).toBe(5)
    c.activate(c.wb.sheets.findIndex((s) => s.name === "Sheet2"))
    c.pivots.setFilter(3, 0, 0, { kind: "label", op: "equal", value: "North" })
    // North is now the only row; its Grand Total cell is D5.
    expect(c.sheet.cells.get(4, 0)?.value).toBe("North")
    c.pivots.showDetails(4, 3)
    expect(c.sheet.name).toBe("Sheet4")
    expect(c.sheet.tables[0].range.r1).toBe(3)
    expect(c.sheet.tables[0].name).toBe("Table2")
  })

  it("PivotTable Options' Enable show details off says so instead", () => {
    const c = setup()
    c.sheet.pivots = [{ ...c.sheet.pivots[0], enableDrill: false }]
    expect(c.pivots.showDetails(4, 1)).toBe(true)
    expect(c.shell.alert).toHaveBeenCalledWith("Cannot change this part of a PivotTable report.")
    expect(c.wb.sheets).toHaveLength(2)
  })
})
