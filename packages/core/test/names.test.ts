import { describe, expect, it } from "vitest"
import { nameProblem, parseCell, parseInput, Workbook, type CellWrite, type Op } from "../src"

function set(wb: Workbook, cells: Record<string, string>) {
  const writes: CellWrite[] = Object.entries(cells).map(([ref, text]) => {
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)!
    return { row, col, data: { value: parsed.value, input: parsed.input } }
  })
  wb.transact([{ kind: "setCells", sheet: 0, cells: writes }], { label: "Typing" })
}
const value = (wb: Workbook, ref: string, sheet = 0) => {
  const { row, col } = parseCell(ref)!
  return wb.sheets[sheet].cells.value(row, col)
}
const run = (wb: Workbook, ...ops: Op[]) => wb.transact(ops, { label: "Test" })

describe("defined names", () => {
  it("a formula uses a name, and recalculates when the name changes or goes", () => {
    const wb = new Workbook(["Sheet1"])
    set(wb, { A1: "2", A2: "3", A3: "10", B1: "=SUM(Data)" })
    expect(value(wb, "B1")).toEqual({ error: "#NAME?" })
    run(wb, { kind: "names", names: [{ name: "Data", ref: "Sheet1!$A$1:$A$2" }] })
    expect(value(wb, "B1")).toBe(5)
    run(wb, { kind: "names", names: [{ name: "Data", ref: "Sheet1!$A$1:$A$3" }] })
    expect(value(wb, "B1")).toBe(15)
    wb.undo()
    expect(value(wb, "B1")).toBe(5)
    wb.undo()
    expect(value(wb, "B1")).toEqual({ error: "#NAME?" })
  })

  it("follows inserted rows and renamed sheets", () => {
    const wb = new Workbook(["Sheet1"])
    set(wb, { A1: "4", B1: "=Rate*2" })
    run(wb, { kind: "names", names: [{ name: "Rate", ref: "Sheet1!$A$1" }] })
    run(wb, { kind: "insertRows", sheet: 0, at: 0, count: 2 })
    expect(wb.names.get("rate")?.ref).toBe("Sheet1!$A$3")
    expect(value(wb, "B3")).toBe(8)
    run(wb, { kind: "renameSheet", sheet: 0, name: "Rates 2026" })
    expect(wb.names.get("Rate")?.ref).toBe("'Rates 2026'!$A$3")
    wb.undo()
    wb.undo()
    expect(wb.names.get("Rate")?.ref).toBe("Sheet1!$A$1")
  })

  it("deleting the sheet a name points at leaves #REF!", () => {
    const wb = new Workbook(["Sheet1", "Sheet2"])
    run(wb, { kind: "names", names: [{ name: "Total", ref: "Sheet2!$B$2" }] })
    run(wb, { kind: "removeSheet", sheet: 1 })
    expect(wb.names.get("Total")?.ref).toBe("#REF!$B$2")
    wb.undo()
    expect(wb.names.get("Total")?.ref).toBe("Sheet2!$B$2")
  })

  it("checks names as Excel does", () => {
    expect(nameProblem("Sales_2026")).toBeNull()
    expect(nameProblem("_tax")).toBeNull()
    expect(nameProblem("A1")).not.toBeNull()
    expect(nameProblem("TAX2026")).not.toBeNull()
    expect(nameProblem("R1C1")).not.toBeNull()
    expect(nameProblem("C")).not.toBeNull()
    expect(nameProblem("1st")).not.toBeNull()
    expect(nameProblem("My Name")).not.toBeNull()
    expect(nameProblem("Rate", [{ name: "RATE", ref: "Sheet1!$A$1" }])).toMatch(/already exists/)
  })
})
