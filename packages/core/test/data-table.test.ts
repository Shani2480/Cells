import { describe, expect, it } from "vitest"
import { breaksDataTable, dataTableBlockAt, dataTableBlocks, dataTableOp, Workbook } from "../src"

type Cell = [row: number, col: number, v: string | number]

function book(cells: Cell[]) {
  const wb = new Workbook()
  set(wb, cells)
  wb.calculate(true)
  return wb
}

function set(wb: Workbook, cells: Cell[]) {
  wb.transact(
    [
      {
        kind: "setCells",
        sheet: 0,
        cells: cells.map(([row, col, v]) => ({
          row,
          col,
          data: typeof v === "string" ? { input: v, value: 0 } : { value: v },
        })),
      },
    ],
    { label: "Set", history: false },
  )
}

const value = (wb: Workbook, row: number, col: number) => wb.sheets[0].cells.value(row, col)

describe("What-If data tables {=TABLE()}", () => {
  it("one input in a column: each row tries its value in the column input cell", () => {
    // A1 rate, A2 = A1 * 100. Table D1:E4: E1 = A2, D2:D4 the rates, E2:E4 =TABLE(,A1).
    const wb = book([
      [0, 0, 0.1],
      [1, 0, "=A1*100"],
      [0, 4, "=A2"],
      [1, 3, 0.2],
      [2, 3, 0.3],
      [3, 3, 0.4],
      [1, 4, "=TABLE(,A1)"],
      [2, 4, "=TABLE(,A1)"],
      [3, 4, "=TABLE(,A1)"],
    ])
    expect([value(wb, 1, 4), value(wb, 2, 4), value(wb, 3, 4)]).toEqual([20, 30, 40])
    // The input cell and the model are untouched.
    expect(value(wb, 0, 0)).toBe(0.1)
    expect(value(wb, 1, 0)).toBe(10)
  })

  it("one input in a row, and through a range (SUM)", () => {
    // A1 x, A2 = SUM(A1:A1, 5). Table: B4 = A2 in the left column, C3:D3 the values, C4:D4 =TABLE(A1,).
    const wb = book([
      [0, 0, 1],
      [1, 0, "=SUM(A1:A1,5)"],
      [3, 1, "=A2"],
      [2, 2, 10],
      [2, 3, 20],
      [3, 2, "=TABLE(A1,)"],
      [3, 3, "=TABLE(A1,)"],
    ])
    expect([value(wb, 3, 2), value(wb, 3, 3)]).toEqual([15, 25])
  })

  it("two inputs: the corner formula with the row and column values", () => {
    // A1 a, A2 b, A3 = A1 * A2. Table at C1: C1 = A3, D1:E1 a values, C2:C3 b values.
    const wb = book([
      [0, 0, 1],
      [1, 0, 1],
      [2, 0, "=A1*A2"],
      [0, 2, "=A3"],
      [0, 3, 2],
      [0, 4, 3],
      [1, 2, 10],
      [2, 2, 100],
      [1, 3, "=TABLE(A1,A2)"],
      [1, 4, "=TABLE(A1,A2)"],
      [2, 3, "=TABLE(A1,A2)"],
      [2, 4, "=TABLE(A1,A2)"],
    ])
    expect([value(wb, 1, 3), value(wb, 1, 4), value(wb, 2, 3), value(wb, 2, 4)]).toEqual([20, 30, 200, 300])
  })

  it("recalculates when the model or the values change", () => {
    const wb = book([
      [0, 0, 1],
      [0, 1, 3],
      [1, 0, "=A1*B1"],
      [0, 4, "=A2"],
      [1, 3, 2],
      [1, 4, "=TABLE(,A1)"],
    ])
    expect(value(wb, 1, 4)).toBe(6)
    set(wb, [[0, 1, 5]])
    expect(value(wb, 1, 4)).toBe(10)
    set(wb, [[1, 3, 7]])
    expect(value(wb, 1, 4)).toBe(35)
    set(wb, [[1, 0, "=A1+B1"]])
    expect(value(wb, 1, 4)).toBe(12)
  })

  it("a formula that doesn't use the input shows its own value; no layout gives #REF!", () => {
    const wb = book([
      [0, 0, 1],
      [0, 4, 42],
      [1, 3, 2],
      [1, 4, "=TABLE(,A1)"],
      [0, 0, 1],
      [3, 0, "=TABLE(,B1)"],
    ])
    expect(value(wb, 1, 4)).toBe(42)
    expect(value(wb, 3, 0)).toEqual({ error: "#REF!" })
  })
})

describe("Data table blocks", () => {
  const table = () =>
    book([
      [0, 0, 1],
      [1, 0, "=A1*2"],
      [0, 2, "=A2"],
      [1, 1, 5],
      [2, 1, 6],
      [1, 2, "=TABLE(,A1)"],
      [2, 2, "=TABLE(,A1)"],
    ])

  it("finds the block of a TABLE cell", () => {
    const s = table().sheets[0]
    expect(dataTableBlockAt(s, 2, 2)).toEqual({ r0: 1, c0: 2, r1: 2, c1: 2 })
    expect(dataTableBlockAt(s, 0, 0)).toBeNull()
    expect(dataTableBlocks(s)).toEqual([{ r0: 1, c0: 2, r1: 2, c1: 2 }])
  })

  it("refuses changing part of a table, allows the whole", () => {
    const s = table().sheets[0]
    const write = (cells: [number, number][]) =>
      breaksDataTable(s, {
        kind: "setCells",
        sheet: 0,
        cells: cells.map(([row, col]) => ({ row, col, data: undefined })),
      })
    expect(write([[1, 2]])).toBe(true)
    expect(
      write([
        [1, 2],
        [2, 2],
      ]),
    ).toBe(false)
    expect(write([[0, 0]])).toBe(false)
    expect(breaksDataTable(s, { kind: "insertRows", sheet: 0, at: 2, count: 1 })).toBe(true)
    expect(breaksDataTable(s, { kind: "insertRows", sheet: 0, at: 1, count: 1 })).toBe(false)
    expect(breaksDataTable(s, { kind: "deleteRows", sheet: 0, at: 2, count: 1 })).toBe(true)
    expect(breaksDataTable(s, { kind: "deleteRows", sheet: 0, at: 0, count: 3 })).toBe(false)
    expect(breaksDataTable(s, { kind: "deleteCols", sheet: 0, at: 0, count: 1 })).toBe(false)
  })

  it("the Data Table command fills the cells past the first row and column", () => {
    const wb = book([
      [0, 0, 1],
      [1, 0, "=A1*2"],
      [0, 2, "=A2"],
      [1, 1, 5],
      [2, 1, 6],
    ])
    const op = dataTableOp(wb.sheets[0], 0, { r0: 0, c0: 1, r1: 2, c1: 2 }, null, "$A$1")!
    wb.transact([op], { label: "Data Table" })
    expect(wb.sheets[0].cells.input(2, 2)).toBe("=TABLE(,$A$1)")
    expect([value(wb, 1, 2), value(wb, 2, 2)]).toEqual([10, 12])
    expect(dataTableOp(wb.sheets[0], 0, { r0: 0, c0: 1, r1: 0, c1: 2 }, null, "A1")).toBeNull()
  })
})
