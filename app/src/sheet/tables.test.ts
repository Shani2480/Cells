import { describe, expect, it, vi } from "vitest"
import { writeInputOp } from "@cells/core"
import { SheetController } from "./controller"
import type { CreateTableInit } from "./tables"
import { CONVERT_QUESTION, OVERLAP_TABLE } from "./tables"

/** A controller with Item/Qty/Price in A1:C4 and the active cell in B2. */
function data() {
  const c = new SheetController()
  const rows = [
    ["Item", "Qty", "Price"],
    ["Pen", "2", "1.5"],
    ["Book", "1", "12"],
    ["Cup", "4", "3"],
  ]
  c.wb.transact(
    rows.flatMap((r, row) => r.map((t, col) => writeInputOp(0, c.sheet, c.wb.styles, row, col, t))),
    { label: "Setup" },
  )
  c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
  c.shell.alert = vi.fn()
  return c
}

const type = (c: SheetController, row: number, col: number, text: string) => {
  c.goTo({ r0: row, c0: col, r1: row, c1: col })
  c.beginEdit("enter", text)
  c.commitEdit()
}

describe("Insert › Table", () => {
  it("offers the data around the active cell, with headers, and makes Table1", () => {
    const c = data()
    let init: CreateTableInit | null = null
    c.shell.createTable = (i) => (init = i)
    c.tables.openCreate()
    expect(init).toEqual({ range: { r0: 0, c0: 0, r1: 3, c1: 2 }, header: true, style: "TableStyleMedium2" })
    expect(c.tables.create(init!.range, init!.header, init!.style)).toBe(true)
    const t = c.sheet.tables[0]
    expect(t).toMatchObject({ name: "Table1", header: true, rowStripes: true, style: "TableStyleMedium2" })
    expect(t.columns.map((x) => x.name)).toEqual(["Item", "Qty", "Price"])
    expect(t.autoFilter?.range).toEqual({ r0: 0, c0: 0, r1: 3, c1: 2 })
    c.wb.undo()
    expect(c.sheet.tables).toEqual([])
  })

  it("without headers, moves the data down and writes Column1, Column2…", () => {
    const c = data()
    c.tables.create({ r0: 1, c0: 0, r1: 3, c1: 1 }, false, "TableStyleLight1")
    expect(c.sheet.cells.get(1, 0)?.value).toBe("Column1")
    expect(c.sheet.cells.get(1, 1)?.value).toBe("Column2")
    expect(c.sheet.cells.get(2, 0)?.value).toBe("Pen")
    expect(c.sheet.cells.get(1, 2)?.value).toBe(1.5)
    expect(c.sheet.tables[0].range).toEqual({ r0: 1, c0: 0, r1: 4, c1: 1 })
  })

  it("refuses to overlap another table", () => {
    const c = data()
    c.tables.create({ r0: 0, c0: 0, r1: 3, c1: 1 }, true, "TableStyleLight1")
    expect(c.tables.create({ r0: 0, c0: 1, r1: 3, c1: 2 }, true, "TableStyleLight1")).toBe(false)
    expect(c.shell.alert).toHaveBeenCalledWith(OVERLAP_TABLE)
  })

  it("a gallery style inside a table restyles it", () => {
    const c = data()
    c.tables.create({ r0: 0, c0: 0, r1: 3, c1: 2 }, true, "TableStyleMedium2")
    c.tables.openCreate("TableStyleDark3")
    expect(c.sheet.tables[0].style).toBe("TableStyleDark3")
  })
})

describe("Table Design", () => {
  const table = () => {
    const c = data()
    c.tables.create({ r0: 0, c0: 0, r1: 3, c1: 2 }, true, "TableStyleMedium2")
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    return c
  }

  it("Total Row: Total and the last column's SUBTOTAL, off again", () => {
    const c = table()
    c.tables.setOption("totals", true)
    expect(c.sheet.tables[0].range.r1).toBe(4)
    expect(c.sheet.cells.get(4, 0)?.value).toBe("Total")
    expect(c.sheet.cells.get(4, 2)?.input).toBe("=SUBTOTAL(109,[Price])")
    expect(c.sheet.cells.get(4, 2)?.value).toBe(16.5)
    c.tables.setOption("totals", false)
    expect(c.sheet.tables[0].range.r1).toBe(3)
    expect(c.sheet.cells.get(4, 0)).toBeUndefined()
  })

  it("style options, rename with formulas following, and names Excel refuses", () => {
    const c = table()
    c.tables.setOption("columnStripes", true)
    c.tables.setOption("filter", false)
    expect(c.sheet.tables[0]).toMatchObject({ columnStripes: true, autoFilter: null })
    type(c, 6, 0, "=SUM(Table1[Qty])")
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    expect(c.tables.rename("Orders")).toBe(true)
    expect(c.sheet.cells.get(6, 0)?.input).toBe("=SUM(Orders[Qty])")
    expect(c.tables.rename("A1")).toBe(false)
    expect(c.sheet.tables[0].name).toBe("Orders")
  })

  it("typing a new header renames the column in formulas", () => {
    const c = table()
    type(c, 6, 0, "=SUM(Table1[Qty])")
    type(c, 0, 1, "Units")
    expect(c.sheet.tables[0].columns[1].name).toBe("Units")
    expect(c.sheet.cells.get(6, 0)?.input).toBe("=SUM(Table1[Units])")
    expect(c.sheet.cells.get(6, 0)?.value).toBe(7)
  })

  it("Resize Table adds a column named from its header", () => {
    const c = table()
    type(c, 0, 3, "Note")
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    expect(c.tables.resize({ r0: 0, c0: 0, r1: 3, c1: 3 })).toBe(true)
    expect(c.sheet.tables[0].columns.map((x) => x.name)).toEqual(["Item", "Qty", "Price", "Note"])
    expect(c.tables.resize({ r0: 1, c0: 0, r1: 3, c1: 3 })).toBe(false)
  })

  it("Convert to Range asks, keeps the look as formats and ends the table", () => {
    const c = table()
    type(c, 6, 0, "=SUM(Table1[Qty])")
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    const confirm = vi.fn((_m: string, ok: () => void) => ok())
    c.shell.confirm = confirm
    c.tables.convertToRange()
    expect(confirm.mock.calls[0][0]).toBe(CONVERT_QUESTION)
    expect(c.sheet.tables).toEqual([])
    expect(c.sheet.cells.get(6, 0)?.input).toBe("=SUM($B$2:$B$4)")
    expect(c.wb.styles.get(c.sheet.styleAt(0, 0))).toMatchObject({ bold: true, fill: "#156082" })
    c.wb.undo()
    expect(c.sheet.tables).toHaveLength(1)
  })

  it("the table's own filter: its dropdowns hide rows, Filter Button off shows them again", () => {
    const c = table()
    c.data.focus = { row: 0, col: 0 }
    expect(c.data.filterColumn(0, { kind: "values", values: ["Pen", "Cup"] })).toBe(true)
    expect([1, 2, 3].map((r) => c.sheet.rows.isHidden(r))).toEqual([false, true, false])
    expect(c.sheet.tables[0].autoFilter?.columns[0]).toBeDefined()
    expect(c.sheet.autoFilter).toBeNull()
    c.data.focus = null
    expect(c.data.records()).toEqual({ shown: 2, total: 3 })
    c.tables.setOption("filter", false)
    expect(c.sheet.rows.isHidden(2)).toBe(false)
    c.wb.undo()
    expect(c.sheet.rows.isHidden(2)).toBe(true)
  })

  it("the total row's dropdown picks a function for the column", () => {
    const c = table()
    c.tables.setOption("totals", true)
    const list = c.cellList(4, 1)!
    expect(list.items).toContain("Count Numbers")
    expect(list.current).toBe("None")
    list.pick("Average")
    expect(c.sheet.cells.get(4, 1)?.input).toBe("=SUBTOTAL(101,[Qty])")
    expect(c.sheet.cells.get(4, 1)?.value).toBeCloseTo(7 / 3)
    expect(c.cellList(4, 1)?.current).toBe("Average")
    c.cellList(4, 0)!.pick("Count")
    expect(c.sheet.cells.get(4, 0)?.value).toBe(3)
    expect(c.cellList(3, 1)).toBeNull()
  })

  it("grows when typing under or beside it; a formula in an empty column fills it", () => {
    const c = table()
    type(c, 0, 3, "Total")
    expect(c.sheet.tables[0].range.c1).toBe(3)
    expect(c.sheet.tables[0].columns[3].name).toBe("Total")
    type(c, 1, 3, "=[@Qty]*[@Price]")
    expect([1, 2, 3].map((r) => c.sheet.cells.get(r, 3)?.value)).toEqual([3, 12, 12])
    expect(c.sheet.tables[0].columns[3].formula).toBe("[@Qty]*[@Price]")
    type(c, 4, 0, "Ink")
    expect(c.sheet.tables[0].range.r1).toBe(4)
    expect(c.sheet.cells.get(4, 3)?.input).toBe("=[@Qty]*[@Price]")
    type(c, 4, 1, "3")
    type(c, 4, 2, "2")
    expect(c.sheet.cells.get(4, 3)?.value).toBe(6)
    type(c, 2, 4, "x")
    expect(c.sheet.tables[0].columns.map((x) => x.name)).toEqual(["Item", "Qty", "Price", "Total", "Column5"])
    expect(c.sheet.cells.get(0, 4)?.value).toBe("Column5")
    c.wb.undo()
    expect(c.sheet.tables[0].range.c1).toBe(3)
    expect(c.sheet.cells.get(0, 4)).toBeUndefined()
  })

  it("Tab wraps inside the table and adds a row at its last cell", () => {
    const c = table()
    c.goTo({ r0: 1, c0: 2, r1: 1, c1: 2 })
    c.keyReady({ key: "Tab", primary: false, shift: false, alt: false })
    expect(c.selection.active).toEqual({ row: 2, col: 0 })
    c.goTo({ r0: 3, c0: 2, r1: 3, c1: 2 })
    c.keyReady({ key: "Tab", primary: false, shift: false, alt: false })
    expect(c.sheet.tables[0].range.r1).toBe(4)
    expect(c.selection.active).toEqual({ row: 4, col: 0 })
    c.keyReady({ key: "Tab", primary: false, shift: true, alt: false })
    expect(c.selection.active).toEqual({ row: 3, col: 2 })
  })

  it("a copied sheet's tables get their own names, and its formulas follow", () => {
    const c = table()
    type(c, 6, 0, "=SUM(Table1[Qty])")
    c.sheets.copy(0, 1)
    const copy = c.wb.sheets[1]
    expect(copy.tables[0].name).toBe("Table13")
    expect(copy.tables[0].id).not.toBe(c.wb.sheets[0].tables[0].id)
    expect(copy.cells.get(6, 0)?.input).toBe("=SUM(Table13[Qty])")
    expect(copy.cells.get(6, 0)?.value).toBe(7)
    expect(c.wb.sheets[0].cells.get(6, 0)?.input).toBe("=SUM(Table1[Qty])")
  })
})
