import { describe, expect, it } from "vitest"
import {
  nextSheetName,
  parseCell,
  parseInput,
  sheetNameProblem,
  snapshotSheet,
  Workbook,
  type CellWrite,
  type Op,
} from "../src"
import { snapshot } from "./helpers"

function set(wb: Workbook, cells: Record<string, string>, sheetIndex = 0) {
  const writes: CellWrite[] = Object.entries(cells).map(([ref, text]) => {
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)
    return { row, col, data: parsed ? { value: parsed.value, input: parsed.input } : undefined }
  })
  wb.transact([{ kind: "setCells", sheet: sheetIndex, cells: writes }], { label: "Typing" })
}

function get(wb: Workbook, ref: string, sheetIndex = 0) {
  const { row, col } = parseCell(ref)!
  const v = wb.sheets[sheetIndex].cells.value(row, col)
  return typeof v === "object" && v !== null && "error" in v ? v.error : v
}

const input = (wb: Workbook, ref: string, sheetIndex = 0) => {
  const { row, col } = parseCell(ref)!
  return wb.sheets[sheetIndex].cells.input(row, col)
}

const run = (wb: Workbook, ...ops: Op[]) => wb.transact(ops, { label: "Sheet" })
const names = (wb: Workbook) => wb.sheets.map((s) => s.name)

describe("sheet operations", () => {
  it("adds a sheet, and undo removes it", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, { kind: "addSheet", sheet: 1, name: "Sheet2" })
    expect(names(wb)).toEqual(["Sheet1", "Sheet2"])
    run(wb, { kind: "addSheet", sheet: 0, name: "Sheet3" })
    expect(names(wb)).toEqual(["Sheet3", "Sheet1", "Sheet2"])
    wb.undo()
    wb.undo()
    expect(names(wb)).toEqual(["Sheet1"])
    wb.redo()
    expect(names(wb)).toEqual(["Sheet1", "Sheet2"])
  })

  it("formulas follow a sheet to its new name, and undo puts the old name back", () => {
    const wb = new Workbook(["Sheet1", "Data"])
    set(wb, { A1: "5" }, 1)
    set(wb, { A1: "=Data!A1*2", A2: "=SUM(Sheet1:Data!A1)", A3: "=data!A1" })
    expect(get(wb, "A1")).toBe(10)
    run(wb, { kind: "renameSheet", sheet: 1, name: "Sales 2026" })
    expect(input(wb, "A1")).toBe("='Sales 2026'!A1*2")
    expect(input(wb, "A2")).toBe("=SUM('Sheet1:Sales 2026'!A1)")
    expect(input(wb, "A3")).toBe("='Sales 2026'!A1")
    expect(get(wb, "A1")).toBe(10)
    set(wb, { A1: "7" }, 1)
    expect(get(wb, "A1")).toBe(14)
    wb.undo()
    wb.undo()
    expect(names(wb)).toEqual(["Sheet1", "Data"])
    expect(input(wb, "A1")).toBe("=Data!A1*2")
    expect(input(wb, "A3")).toBe("=data!A1")
    expect(get(wb, "A1")).toBe(10)
  })

  it("deleting a sheet turns references to it into #REF!, and undo brings back the sheet and the formulas", () => {
    const wb = new Workbook(["Sheet1", "Data", "Sheet3"])
    set(wb, { B2: "4", C3: "x" }, 1)
    wb.transact([{ kind: "format", sheet: 1, ranges: [{ r0: 1, r1: 1, c0: 1, c1: 1 }], patch: { bold: true } }], {
      label: "Bold",
    })
    wb.transact([{ kind: "axis", sheet: 1, axis: "cols", entries: [{ index: 2, size: 120, hidden: false }] }], {
      label: "Width",
    })
    wb.transact([{ kind: "merge", sheet: 1, ranges: [{ r0: 5, r1: 6, c0: 0, c1: 2 }] }], { label: "Merge" })
    run(wb, { kind: "sheetProps", sheet: 1, tabColor: "#C00000" })
    set(wb, { A1: "=Data!B2+1" })
    set(wb, { A1: "=Data!B2*3" }, 2)
    const before = snapshot(wb.sheets[1])
    expect(get(wb, "A1")).toBe(5)

    run(wb, { kind: "removeSheet", sheet: 1 })
    expect(names(wb)).toEqual(["Sheet1", "Sheet3"])
    expect(input(wb, "A1")).toBe("=#REF!B2+1")
    expect(get(wb, "A1")).toBe("#REF!")
    expect(input(wb, "A1", 1)).toBe("=#REF!B2*3")

    wb.undo()
    expect(names(wb)).toEqual(["Sheet1", "Data", "Sheet3"])
    expect(snapshot(wb.sheets[1])).toEqual(before)
    expect(wb.sheets[1].tabColor).toBe("#C00000")
    expect(input(wb, "A1")).toBe("=Data!B2+1")
    expect(input(wb, "A1", 2)).toBe("=Data!B2*3")
    expect(get(wb, "A1")).toBe(5)
    expect(get(wb, "A1", 2)).toBe(12)
  })

  it("moves a sheet; formulas keep pointing at it by name", () => {
    const wb = new Workbook(["A", "B", "C"])
    set(wb, { A1: "3" }, 2)
    set(wb, { A1: "=C!A1+1" }, 0)
    run(wb, { kind: "moveSheet", sheet: 2, to: 0 })
    expect(names(wb)).toEqual(["C", "A", "B"])
    expect(get(wb, "A1", 1)).toBe(4)
    set(wb, { A1: "10" }, 0)
    expect(get(wb, "A1", 1)).toBe(11)
    wb.undo()
    wb.undo()
    expect(names(wb)).toEqual(["A", "B", "C"])
    expect(get(wb, "A1", 0)).toBe(4)
  })

  it("tab colour and hiding undo", () => {
    const wb = new Workbook(["Sheet1", "Sheet2"])
    run(wb, { kind: "sheetProps", sheet: 1, tabColor: "#00B050", hidden: true })
    expect(wb.sheets[1].tabColor).toBe("#00B050")
    expect(wb.sheets[1].hidden).toBe(true)
    run(wb, { kind: "sheetProps", sheet: 1, tabColor: null })
    expect(wb.sheets[1].tabColor).toBeUndefined()
    wb.undo()
    expect(wb.sheets[1].tabColor).toBe("#00B050")
    wb.undo()
    expect(wb.sheets[1].tabColor).toBeUndefined()
    expect(wb.sheets[1].hidden).toBe(false)
  })

  it("a background picture undoes and travels with a copied sheet", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, { kind: "sheetProps", sheet: 0, background: "data:image/png;base64,AA==" })
    expect(wb.sheets[0].background).toBe("data:image/png;base64,AA==")
    run(wb, { kind: "addSheet", sheet: 1, name: "Copy", snapshot: snapshotSheet(wb.sheets[0]) })
    expect(wb.sheets[1].background).toBe("data:image/png;base64,AA==")
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].background).toBeUndefined()
  })

  it("checks names as Excel does", () => {
    expect(sheetNameProblem("Sales", ["Sheet1"])).toBeNull()
    expect(sheetNameProblem("sheet1", ["Sheet1"])).toMatch(/already taken/)
    expect(sheetNameProblem("", [])).toMatch(/invalid/)
    expect(sheetNameProblem("a/b", [])).toMatch(/invalid/)
    expect(sheetNameProblem("x".repeat(32), [])).toMatch(/invalid/)
    expect(sheetNameProblem("'q", [])).toMatch(/invalid/)
    expect(sheetNameProblem("History", [])).toMatch(/reserved/)
    expect(nextSheetName(["Sheet1"])).toBe("Sheet2")
    expect(nextSheetName(["Sheet1", "Sheet3"])).toBe("Sheet4")
    expect(nextSheetName(["Sheet1", "Data", "Sheet3"])).toBe("Sheet4")
  })
})
