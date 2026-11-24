import { describe, expect, it, vi } from "vitest"
import { pivotOptions, withFieldIn, withoutField, Workbook, type Selection } from "@cells/core"
import { SheetController } from "./controller"
import { PIVOT_NO_DATA_MESSAGE } from "./pivots"

/** Product, Region, Amount on Data!A1:C7 (or `many` products); Sum of Amount by Region and Product at E3. */
function setup(many = 0) {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.confirm = vi.fn((_, yes: () => void) => yes())
  const rows: (string | number)[][] = [
    ["Product", "Region", "Amount"],
    ["Pens", "East", 10],
    ["Ink", "West", 40],
    ["Paper", "East", 25],
    ["Pens", "West", 5],
    ["Clips", "East", 1],
    ["Ink", "East", 2],
  ]
  for (let i = 0; i < many; i++) rows.push([`Item ${i + 1}`, "East", i])
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  c.pivots.create(`A1:C${rows.length}`, { kind: "existing", sheet: 0, at: { row: 2, col: 4 } }, undefined, {
    title: "",
    rows: [1, 0],
    cols: [],
    values: [2],
  })
  c.wb.clearHistory()
  c.goTo({ r0: 3, c0: 4, r1: 3, c1: 4 })
  return c
}
const options = (c: SheetController) => {
  const at = c.pivots.activeWithCache()!
  return pivotOptions(at.table, at.cache)
}
const merges = (c: SheetController) => c.sheet.merges.all.map((m) => `${m.r0},${m.c0}:${m.r1},${m.c1}`).sort()

describe("PivotTable Options", () => {
  it("Classic layout goes tabular; merged labels merge and center on the sheet, as one undo step", () => {
    const c = setup()
    expect(c.pivots.setOptions({ ...options(c), gridDropZones: true, mergeLabels: true })).toBeNull()
    // East over its four products (rows 4–7), its total across both label columns, West, its total, Grand Total.
    expect(merges(c)).toEqual(["3,4:6,4", "7,4:7,5", "8,4:9,4", "10,4:10,5", "11,4:11,5"].sort())
    const east = c.wb.styles.get(c.sheet.cells.style(3, 4))
    expect([c.sheet.cells.get(3, 4)?.value, east.hAlign, east.vAlign]).toEqual(["East", "center", "center"])
    c.wb.undo()
    expect(merges(c)).toEqual([])
    expect(c.sheet.cells.get(3, 4)?.value).toBe("East")
    expect(options(c).gridDropZones).toBe(false)
    c.wb.redo()
    // Turned off, the labels come apart and lose the centring (Excel's).
    c.goTo({ r0: 3, c0: 4, r1: 3, c1: 4 })
    expect(c.pivots.setOptions({ ...options(c), mergeLabels: false })).toBeNull()
    expect(merges(c)).toEqual([])
    expect(c.wb.styles.get(c.sheet.cells.style(3, 4)).hAlign).toBeUndefined()
  })

  it("lays the report filters out over then down, and the filters still work from their new cells", () => {
    const c = setup()
    c.pivots.change((t, k) => withFieldIn(withoutField(t, 0, "rows"), k, 0, "pages"))
    const before = c.pivots.activeWithCache()!.table
    expect(before.location.r0).toBe(2)
    c.goTo({ r0: 2, c0: 4, r1: 2, c1: 4 })
    expect(c.pivots.setOptions({ ...options(c), pageOverThenDown: true, pageWrap: 1 })).toBeNull()
    // One filter: the same place either way.
    expect(c.sheet.cells.get(0, 4)?.value).toBe("Product")
    expect(c.sheet.cells.get(0, 5)?.value).toBe("(All)")
    expect(c.pivots.choices(0, 5, 0)?.name).toBe("Product")
  })

  it("keeps the cache's options: no source data saved, items retained", () => {
    const c = setup()
    expect(c.pivots.setOptions({ ...options(c), saveData: false, missingItems: "none" })).toBeNull()
    expect(c.wb.pivotCaches[0]).toMatchObject({ saveData: false, missingItemsLimit: 0 })
    expect(options(c)).toMatchObject({ saveData: false, missingItems: "none" })
    // As read from such a file: no records, so changes wait for a refresh (Excel's message).
    c.wb.pivotCaches = c.wb.pivotCaches.map((k) => ({ ...k, records: [] }))
    c.goTo({ r0: 3, c0: 4, r1: 3, c1: 4 })
    expect(c.pivots.setOptions({ ...options(c), mergeLabels: true })).toBeNull()
    expect(c.shell.alert).toHaveBeenCalledWith(PIVOT_NO_DATA_MESSAGE)
    expect(c.pivots.refresh()).toBe(true)
    expect(c.wb.pivotCaches[0].records.length).toBe(6)
  })

  it("Set print titles prints the PivotTable's header rows and label columns on every page", () => {
    const c = setup(120)
    expect(c.pivots.setOptions({ ...options(c), fieldPrintTitles: true })).toBeNull()
    const pages = c.print.pages().pages
    expect(pages.length).toBeGreaterThan(1)
    expect(pages[1].titleRows).toEqual([2, 2])
    c.print.pages(c.sheet, {}, { ...c.sheet.pageSetup, titleRows: [0, 0] }).pages.forEach((p, i) => {
      if (i) expect(p.titleRows).toEqual([0, 0])
    })
  })
})
