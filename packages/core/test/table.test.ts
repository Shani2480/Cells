// Tables and structured references. Results as Excel 365 gives them for the same table (made with
// ListObjects.Add on A1:C5 of the sample below).
import { describe, expect, it } from "vitest"
import {
  DEFAULT_TABLE_STYLE,
  parseStructured,
  printStructured,
  tableCellLook,
  Workbook,
  writeInputOp,
  type Op,
  type Table,
} from "../src"

function sample() {
  const wb = new Workbook(["Data", "Other"])
  const rows = [
    ["Item", "Qty", "Price"],
    ["Pen", "2", "1.5"],
    ["Book", "1", "12"],
    ["Cup", "4", "3"],
  ]
  const ops: Op[] = []
  rows.forEach((r, row) => r.forEach((t, col) => ops.push(writeInputOp(0, wb.sheets[0], wb.styles, row, col, t))))
  const table: Table = {
    id: 1,
    name: "Sales",
    range: { r0: 0, c0: 0, r1: 4, c1: 2 },
    header: true,
    totals: true,
    columns: [{ name: "Item", totalsLabel: "Total" }, { name: "Qty", totals: "sum" }, { name: "Price" }],
    style: DEFAULT_TABLE_STYLE,
    firstColumn: false,
    lastColumn: false,
    rowStripes: true,
    columnStripes: false,
    autoFilter: null,
  }
  ops.push({ kind: "tables", sheet: 0, tables: [table] })
  wb.transact(ops, { label: "Setup" })
  return wb
}

const type = (wb: Workbook, sheet: number, row: number, col: number, text: string) =>
  wb.transact([writeInputOp(sheet, wb.sheets[sheet], wb.styles, row, col, text)], { label: "Typing" })
const value = (wb: Workbook, sheet: number, row: number, col: number) => wb.sheets[sheet].cells.get(row, col)?.value

describe("structured references", () => {
  it("read and write Excel's forms", () => {
    expect(parseStructured("Sales[Qty]")).toMatchObject({ table: "Sales", data: true, c0: "Qty", c1: "Qty" })
    expect(parseStructured("Sales[]")).toMatchObject({ data: true, headers: false, c0: null })
    expect(parseStructured("Sales[#All]")).toMatchObject({ headers: true, data: true, totals: true })
    expect(parseStructured("[@Qty]")).toMatchObject({ table: null, thisRow: true, c0: "Qty" })
    expect(parseStructured("Sales[@[Unit Price]]")).toMatchObject({ thisRow: true, c0: "Unit Price" })
    expect(parseStructured("Sales[[#Headers],[#Data],[Qty]:[Price]]")).toMatchObject({
      headers: true,
      data: true,
      totals: false,
      c0: "Qty",
      c1: "Price",
    })
    expect(parseStructured("Sales[[#This Row],[Qty]]")).toMatchObject({ thisRow: true, c0: "Qty" })
    expect(parseStructured("Sales[['#Hash]]")).toMatchObject({ c0: "#Hash" })
    expect(parseStructured("Sales[Qty")).toBeNull()

    const thisRow = parseStructured("Sales[[#This Row],[Unit Price]]")!
    expect(printStructured(thisRow, "Sales", true)).toBe("[@[Unit Price]]")
    expect(printStructured(thisRow, "Sales", false)).toBe("Sales[@[Unit Price]]")
    expect(printStructured(thisRow, "Sales", true, true)).toBe("Sales[[#This Row],[Unit Price]]")
    expect(printStructured(parseStructured("[@Qty]")!, "Sales", true)).toBe("[@Qty]")
    expect(printStructured(parseStructured("Sales[[#Totals],[Qty]]")!, "Sales", false)).toBe("Sales[[#Totals],[Qty]]")
    expect(printStructured(parseStructured("Sales[]")!, "Sales", false)).toBe("Sales[#Data]")
    expect(printStructured(parseStructured("Sales[Qty]")!, "Sales", true)).toBe("[Qty]")
  })

  it("calculate as Excel does", () => {
    const wb = sample()
    type(wb, 1, 0, 0, "=SUM(Sales[Qty])")
    type(wb, 1, 1, 0, "=COUNTA(Sales[#All])")
    type(wb, 1, 2, 0, "=ROWS(Sales[[#Headers],[#Data],[Item]])")
    type(wb, 1, 3, 0, "=Sales[[#Headers],[Price]]")
    type(wb, 1, 4, 0, "=SUMPRODUCT(Sales[Qty],Sales[Price])")
    type(wb, 1, 5, 0, "=Sales[[#Totals],[Item]]")
    type(wb, 1, 6, 0, "=SUM(Nope[Qty])")
    type(wb, 1, 7, 0, "=SUM(Sales[Nope])")
    expect([0, 1, 2, 3, 4, 5].map((r) => value(wb, 1, r, 0))).toEqual([7, 12, 4, "Price", 27, 0])
    expect(value(wb, 1, 6, 0)).toEqual({ error: "#REF!" })
    expect(value(wb, 1, 7, 0)).toEqual({ error: "#REF!" })
  })

  it("[@Column] reads the formula's own row; outside the data rows it is #VALUE!", () => {
    const wb = sample()
    type(wb, 0, 1, 3, "=Sales[@Qty]*Sales[@Price]")
    type(wb, 0, 4, 1, "=SUBTOTAL(109,[Qty])")
    expect(value(wb, 0, 1, 3)).toBe(3)
    expect(value(wb, 0, 4, 1)).toBe(7)
    type(wb, 0, 2, 4, "=Sales[@Qty]")
    expect(value(wb, 0, 2, 4)).toBe(1)
    type(wb, 0, 7, 4, "=Sales[@Qty]")
    expect(value(wb, 0, 7, 4)).toEqual({ error: "#VALUE!" })
  })

  it("follow changes: a new value, rows inserted inside the table, a renamed header", () => {
    const wb = sample()
    type(wb, 1, 0, 0, "=SUM(Sales[Qty])")
    type(wb, 0, 2, 1, "5")
    expect(value(wb, 1, 0, 0)).toBe(11)
    wb.transact([{ kind: "insertRows", sheet: 0, at: 2, count: 1 }], { label: "Insert" })
    expect(wb.sheets[0].tables[0].range).toEqual({ r0: 0, c0: 0, r1: 5, c1: 2 })
    type(wb, 0, 2, 1, "10")
    expect(value(wb, 1, 0, 0)).toBe(21)
    type(wb, 0, 0, 1, "Units")
    expect(wb.sheets[0].tables[0].columns[1].name).toBe("Units")
    expect(value(wb, 1, 0, 0)).toEqual({ error: "#REF!" })
    wb.undo()
    expect(wb.sheets[0].tables[0].columns[1].name).toBe("Qty")
    expect(value(wb, 1, 0, 0)).toBe(21)
  })

  it("inserted columns become table columns; deleting the table's rows drops it", () => {
    const wb = sample()
    wb.transact([{ kind: "insertCols", sheet: 0, at: 1, count: 1 }], { label: "Insert" })
    expect(wb.sheets[0].tables[0].columns.map((c) => c.name)).toEqual(["Item", "Column4", "Qty", "Price"])
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 0, count: 5 }], { label: "Delete" })
    expect(wb.sheets[0].tables).toEqual([])
    wb.undo()
    expect(wb.sheets[0].tables[0].range).toEqual({ r0: 0, c0: 0, r1: 4, c1: 3 })
  })
})

describe("table styles", () => {
  it("TableStyleMedium2 as Excel draws it: blue header, banded rows, total row line", () => {
    const t = sample().sheets[0].tables[0]
    expect(tableCellLook(t, 0, 0)).toMatchObject({ fill: "#156082", bold: true, fontColor: "#FFFFFF" })
    expect(tableCellLook(t, 1, 1)).toMatchObject({ fill: "#C0E6F5" })
    expect(tableCellLook(t, 2, 1)?.fill).toBeUndefined()
    expect(tableCellLook(t, 4, 1)).toMatchObject({ bold: true, borderTop: { style: "double", color: "#156082" } })
    expect(tableCellLook({ ...t, rowStripes: false }, 1, 1)?.fill).toBeUndefined()
    expect(tableCellLook({ ...t, style: null }, 0, 0)).toBeNull()
  })
})

describe("formulas follow the table", () => {
  it("renaming the table or a column rewrites every reference to it", async () => {
    const { renameTableOps } = await import("../src")
    const wb = sample()
    type(wb, 1, 0, 0, "=SUM(Sales[Qty])+Sales[[#Totals],[Qty]]")
    type(wb, 0, 4, 2, "=SUBTOTAL(109,[Price])")
    const t = wb.sheets[0].tables[0]
    wb.transact(renameTableOps(wb.sheets, t, "Orders", { from: "Price", to: "Cost" }), { label: "Rename" })
    expect(wb.sheets[1].cells.get(0, 0)?.input).toBe("=SUM(Orders[Qty])+Orders[[#Totals],[Qty]]")
    expect(wb.sheets[0].cells.get(4, 2)?.input).toBe("=SUBTOTAL(109,[Cost])")
  })

  it("Convert to Range turns references into ordinary ones", async () => {
    const { unstructureOps } = await import("../src")
    const wb = sample()
    type(wb, 1, 0, 0, "=SUM(Sales[Qty])")
    type(wb, 0, 2, 3, "=Sales[@Qty]*2")
    type(wb, 0, 4, 1, "=SUBTOTAL(109,[Qty])")
    wb.transact(unstructureOps(wb.sheets, wb.sheets[0].tables[0]), { label: "Convert" })
    expect(wb.sheets[1].cells.get(0, 0)?.input).toBe("=SUM(Data!$B$2:$B$4)")
    expect(wb.sheets[0].cells.get(2, 3)?.input).toBe("=B3*2")
    expect(wb.sheets[0].cells.get(4, 1)?.input).toBe("=SUBTOTAL(109,$B$2:$B$4)")
    expect(value(wb, 1, 0, 0)).toBe(7)
  })
})
