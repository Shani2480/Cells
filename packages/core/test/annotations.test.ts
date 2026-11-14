import { describe, expect, it } from "vitest"
import { DEFAULT_PROTECTION, hashPassword, protectionBlocks, Workbook, type CellNote, type Op } from "../src"

const note = (text: string): CellNote => ({ kind: "note", author: "Alex", text, shown: false })
const run = (wb: Workbook, ...ops: Op[]) => wb.transact(ops, { label: "Test" })
const at = (wb: Workbook, row: number, col: number) => wb.sheets[0].annotations.get(row, col)

describe("notes and comments", () => {
  it("sets a note, and undo takes it away", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, { kind: "annotations", sheet: 0, writes: [{ row: 1, col: 1, value: note("Check") }] })
    expect(at(wb, 1, 1)).toEqual(note("Check"))
    wb.undo()
    expect(at(wb, 1, 1)).toBeUndefined()
    wb.redo()
    expect(at(wb, 1, 1)).toEqual(note("Check"))
  })

  it("moves with inserted rows and goes with deleted ones, and undo brings it back", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, {
      kind: "annotations",
      sheet: 0,
      writes: [
        { row: 4, col: 0, value: note("a") },
        { row: 9, col: 2, value: note("b") },
      ],
    })
    run(wb, { kind: "insertRows", sheet: 0, at: 2, count: 3 })
    expect(at(wb, 7, 0)).toEqual(note("a"))
    expect(at(wb, 12, 2)).toEqual(note("b"))
    run(wb, { kind: "deleteCols", sheet: 0, at: 2, count: 1 })
    expect(at(wb, 12, 2)).toBeUndefined()
    wb.undo()
    expect(at(wb, 12, 2)).toEqual(note("b"))
    wb.undo()
    expect(at(wb, 4, 0)).toEqual(note("a"))
  })

  it("lists annotations in reading order", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, {
      kind: "annotations",
      sheet: 0,
      writes: [
        { row: 3, col: 0, value: note("c") },
        { row: 0, col: 5, value: note("b") },
        { row: 0, col: 1, value: note("a") },
      ],
    })
    expect(wb.sheets[0].annotations.all().map((a) => (a.value as CellNote).text)).toEqual(["a", "b", "c"])
  })

  it("travels with a copied sheet", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, { kind: "annotations", sheet: 0, writes: [{ row: 0, col: 0, value: note("x") }] })
    run(wb, { kind: "removeSheet", sheet: 0 }, { kind: "addSheet", sheet: 0, name: "Sheet1" })
    expect(at(wb, 0, 0)).toBeUndefined()
    wb.undo()
    expect(at(wb, 0, 0)).toEqual(note("x"))
  })
})

describe("sheet protection", () => {
  it("hashes passwords as Excel's legacy sheet hash", () => {
    expect(hashPassword("secret")).toBe("DAA7")
    expect(hashPassword("")).toBe("CE4B")
  })

  it("protects and unprotects with undo", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, {
      kind: "protect",
      sheet: 0,
      protection: {
        selectLocked: true,
        selectUnlocked: true,
        formatCells: true,
        formatColumns: false,
        formatRows: false,
        insertColumns: false,
        insertRows: false,
        insertHyperlinks: false,
        deleteColumns: false,
        deleteRows: false,
        sort: false,
        autoFilter: false,
        pivotTables: false,
        editObjects: false,
        editScenarios: false,
      },
    })
    expect(wb.sheets[0].protection?.formatCells).toBe(true)
    wb.undo()
    expect(wb.sheets[0].protection).toBeNull()
  })
})

describe("protection checks", () => {
  const protect = (wb: Workbook, extra: Partial<import("../src").SheetProtection> = {}) =>
    run(wb, { kind: "protect", sheet: 0, protection: { ...DEFAULT_PROTECTION, ...extra } })

  it("refuses typing in locked cells and allows it in unlocked ones", () => {
    const wb = new Workbook(["Sheet1"])
    run(wb, { kind: "format", sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], patch: { locked: false } })
    protect(wb)
    const write = (row: number): Op => ({ kind: "setCells", sheet: 0, cells: [{ row, col: 0, data: { value: 1 } }] })
    expect(protectionBlocks(wb.sheets[0], wb.styles, write(0))).toBe(false)
    expect(protectionBlocks(wb.sheets[0], wb.styles, write(1))).toBe(true)
  })

  it("allows formatting and row inserts only when the options say so", () => {
    const wb = new Workbook(["Sheet1"])
    protect(wb, { insertRows: true })
    const format: Op = { kind: "format", sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], patch: { bold: true } }
    expect(protectionBlocks(wb.sheets[0], wb.styles, format)).toBe(true)
    expect(protectionBlocks(wb.sheets[0], wb.styles, { kind: "insertRows", sheet: 0, at: 1, count: 1 })).toBe(false)
    expect(protectionBlocks(wb.sheets[0], wb.styles, { kind: "insertCols", sheet: 0, at: 1, count: 1 })).toBe(true)
  })

  it("checks a big unlocked column quickly by its column format", () => {
    const wb = new Workbook(["Sheet1"])
    const col = { r0: 0, c0: 2, r1: 1_048_575, c1: 2 }
    run(wb, { kind: "format", sheet: 0, ranges: [col], patch: { locked: false } })
    protect(wb)
    const move: Op = { kind: "moveCells", sheet: 0, source: { ...col, r1: 999_999 }, dr: 1, dc: 0 }
    expect(protectionBlocks(wb.sheets[0], wb.styles, move)).toBe(false)
  })
})
