import { describe, expect, it } from "vitest"
import {
  copyFormula,
  fillOp,
  moveBlockOps,
  parseCell,
  parseInput,
  pasteBlockOp,
  readBlock,
  rewriteFormula,
  shiftMapper,
  Workbook,
  type CellWrite,
  type Op,
} from "../src"

function book(cells: Record<string, string>, sheets = ["Sheet1"]) {
  const wb = new Workbook(sheets)
  for (const [key, text] of Object.entries(cells)) {
    const [sheetPart, ref] = key.includes("!") ? key.split("!") : ["0", key]
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)!
    const writes: CellWrite[] = [{ row, col, data: { value: parsed.value, input: parsed.input } }]
    wb.transact([{ kind: "setCells", sheet: Number(sheetPart), cells: writes }], { label: "Typing" })
  }
  return wb
}

function input(wb: Workbook, ref: string, sheet = 0) {
  const { row, col } = parseCell(ref)!
  return wb.sheets[sheet].cells.get(row, col)?.input
}

function value(wb: Workbook, ref: string, sheet = 0) {
  const { row, col } = parseCell(ref)!
  const v = wb.sheets[sheet].cells.value(row, col)
  return typeof v === "object" && v !== null && "error" in v ? v.error : v
}

const run = (wb: Workbook, ...ops: Op[]) => wb.transact(ops, { label: "Edit" })
const range = (r0: number, c0: number, r1: number, c1: number) => ({ r0, c0, r1, c1 })

describe("copying formulas", () => {
  it.each([
    ["=A1+$B$1+$C2+D$3", 1, 1, "=B2+$B$1+$C3+E$3"],
    ["=SUM(A1:B2)", 2, 0, "=SUM(A3:B4)"],
    ["=SUM(A:A)", 5, 1, "=SUM(B:B)"],
    ["=SUM(1:1)", 1, 5, "=SUM(2:2)"],
    ["=A1", -1, 0, "=#REF!"],
    ["=Sheet2!A1 + 'My Sheet'!B2", 1, 0, "=Sheet2!A2 + 'My Sheet'!B3"],
    ['=INDIRECT("A1")&A1', 1, 0, '=INDIRECT("A1")&A2'],
    ["=XFD1", 0, 1, "=#REF!"],
  ])("%s moved by (%i, %i) is %s", (formula, dr, dc, expected) => {
    expect(copyFormula(formula, dr, dc)).toBe(expected)
  })
})

describe("reference shifting", () => {
  const rows = (formula: string, at: number, count: number) =>
    rewriteFormula(formula, "Sheet1", "Sheet1", shiftMapper("rows", at, count)) ?? formula

  it.each([
    // Inserting two rows at row 5 (index 4)
    ["=A4+A5+A6", 4, 2, "=A4+A7+A8"],
    ["=SUM(A1:A10)", 4, 2, "=SUM(A1:A12)"],
    ["=SUM(A5:A10)", 4, 2, "=SUM(A7:A12)"],
    ["=SUM(A1:A4)", 4, 2, "=SUM(A1:A4)"],
    ["=SUM(B:B)", 4, 2, "=SUM(B:B)"],
    ["=SUM(5:6)", 4, 2, "=SUM(7:8)"],
    ["=$A$5", 4, 2, "=$A$7"],
    // Deleting rows 5-6
    ["=A5", 4, -2, "=#REF!"],
    ["=A4+A7", 4, -2, "=A4+A5"],
    ["=SUM(A1:A10)", 4, -2, "=SUM(A1:A8)"],
    ["=SUM(A5:A6)", 4, -2, "=SUM(#REF!)"],
    ["=SUM(A5:A10)", 4, -2, "=SUM(A5:A8)"],
    ["=SUM(A2:A5)", 4, -2, "=SUM(A2:A4)"],
    ["=SUM(A10:A1)", 4, -2, "=SUM(A8:A1)"],
  ])("%s at row index %i by %i is %s", (formula, at, count, expected) => {
    expect(rows(formula, at, count)).toBe(expected)
  })

  it("only touches references to the changed sheet", () => {
    const map = shiftMapper("rows", 0, 1)
    expect(rewriteFormula("=A1+Sheet2!A1", "Sheet1", "Sheet2", map)).toBe("=A1+Sheet2!A2")
    expect(rewriteFormula("=A1+sheet2!A1", "Sheet2", "Sheet2", map)).toBe("=A2+sheet2!A2")
    expect(rewriteFormula("=Sheet3!A1", "Sheet1", "Sheet2", map)).toBeNull()
  })

  it("keeps spacing and spelling", () => {
    expect(rewriteFormula("= SUM( A1 , B1 )", "S", "S", shiftMapper("cols", 0, 1))).toBe("= SUM( B1 , C1 )")
  })

  it("moves cells of a band only (Insert Cells)", () => {
    const map = shiftMapper("rows", 2, 1, [0, 1]) // shift A3:B3 down
    expect(rewriteFormula("=A5+B5+C5", "S", "S", map)).toBe("=A6+B6+C5")
    expect(rewriteFormula("=SUM(A3:C9)", "S", "S", map)).toBeNull()
  })
})

describe("structural edits in a workbook", () => {
  it("insert and delete rows rewrite formulas everywhere, and undo restores them", () => {
    const wb = book({ A1: "1", A2: "2", A3: "3", B1: "=SUM(A1:A3)", "1!A1": "=Sheet1!A3*10" }, ["Sheet1", "Sheet2"])
    expect(value(wb, "B1")).toBe(6)
    run(wb, { kind: "insertRows", sheet: 0, at: 1, count: 2 })
    expect(input(wb, "B1")).toBe("=SUM(A1:A5)")
    expect(input(wb, "A1", 1)).toBe("=Sheet1!A5*10")
    expect(value(wb, "A1", 1)).toBe(30)
    wb.undo()
    expect(input(wb, "B1")).toBe("=SUM(A1:A3)")
    expect(input(wb, "A1", 1)).toBe("=Sheet1!A3*10")
    run(wb, { kind: "deleteRows", sheet: 0, at: 2, count: 1 })
    expect(input(wb, "B1")).toBe("=SUM(A1:A2)")
    expect(value(wb, "B1")).toBe(3)
    expect(input(wb, "A1", 1)).toBe("=Sheet1!#REF!*10")
    expect(value(wb, "A1", 1)).toBe("#REF!")
    wb.undo()
    expect(input(wb, "A1", 1)).toBe("=Sheet1!A3*10")
    expect(value(wb, "A1", 1)).toBe(30)
    wb.redo()
    expect(input(wb, "A1", 1)).toBe("=Sheet1!#REF!*10")
    wb.undo()
    expect(value(wb, "B1")).toBe(6)
  })

  it("a formula below the change moves and keeps pointing at its cells", () => {
    const wb = book({ A1: "5", A10: "=A1*2" })
    run(wb, { kind: "deleteCols", sheet: 0, at: 1, count: 3 })
    run(wb, { kind: "insertRows", sheet: 0, at: 0, count: 1 })
    expect(input(wb, "A11")).toBe("=A2*2")
    expect(value(wb, "A11")).toBe(10)
  })

  it("delete cells shifting up turns references to them into #REF!", () => {
    const wb = book({ A1: "1", A2: "2", A3: "3", C1: "=A2+A3" })
    run(wb, { kind: "deleteCells", sheet: 0, range: range(1, 0, 1, 0), shift: "up" })
    expect(input(wb, "C1")).toBe("=#REF!+A2")
    wb.undo()
    expect(input(wb, "C1")).toBe("=A2+A3")
    expect(value(wb, "C1")).toBe(5)
  })
})

describe("copy, fill and move", () => {
  it("paste adjusts relative references", () => {
    const wb = book({ A1: "1", A2: "2", B1: "=A1*$A$1" })
    const block = readBlock(wb.sheets[0], range(0, 1, 0, 1))
    run(wb, pasteBlockOp(0, range(1, 1, 1, 1), block))
    expect(input(wb, "B2")).toBe("=A2*$A$1")
    expect(value(wb, "B2")).toBe(2)
  })

  it("the fill handle copies formulas along", () => {
    const wb = book({ A1: "1", A2: "2", A3: "3", B1: "=A1+1" })
    run(wb, fillOp(0, wb.sheets[0], range(0, 1, 0, 1), range(0, 1, 2, 1))!)
    expect(input(wb, "B3")).toBe("=A3+1")
    expect(value(wb, "B3")).toBe(4)
  })

  it("cut and paste moves references with the cells", () => {
    const wb = book({ A1: "7", B1: "=A1", C1: "=SUM(A1:A1)", D5: "=D1" })
    run(wb, ...moveBlockOps(0, range(0, 0, 0, 0), range(4, 3, 4, 3)))
    expect(input(wb, "B1")).toBe("=D5")
    expect(input(wb, "C1")).toBe("=SUM(D5:D5)")
    expect(value(wb, "B1")).toBe(7)
    expect(value(wb, "A1")).toBeUndefined()
    wb.undo()
    expect(input(wb, "B1")).toBe("=A1")
    expect(input(wb, "D5")).toBe("=D1")
    expect(value(wb, "A1")).toBe(7)
  })

  it("a moved formula keeps its references to cells outside the move", () => {
    const wb = book({ A1: "3", B1: "=A1*2" })
    run(wb, ...moveBlockOps(0, range(0, 1, 0, 1), range(5, 5, 5, 5)))
    expect(input(wb, "F6")).toBe("=A1*2")
    expect(value(wb, "F6")).toBe(6)
  })
})
