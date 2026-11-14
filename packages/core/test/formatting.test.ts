import { describe, expect, it } from "vitest"
import { MAX_COLS, MAX_ROWS, parseRange } from "../src/address"
import { writeInputOp } from "../src/edit"
import type { Op } from "../src/ops"
import { CLEAR_ALL, EXPLICIT_DEFAULT_ID, type StylePatch } from "../src/styles"
import { Workbook } from "../src/workbook"

const range = (text: string) => {
  if (/^[A-Z]+:[A-Z]+$/.test(text)) {
    const [a, b] = text.split(":").map((c) => parseRange(`${c}1`)!.c0)
    return { r0: 0, r1: MAX_ROWS - 1, c0: a, c1: b }
  }
  if (/^\d+:\d+$/.test(text)) {
    const [a, b] = text.split(":").map((r) => Number(r) - 1)
    return { r0: a, r1: b, c0: 0, c1: MAX_COLS - 1 }
  }
  return parseRange(text)!
}

function setup() {
  const wb = new Workbook(["Sheet1"])
  const sheet = wb.sheets[0]
  const run = (...ops: Op[]) => wb.transact(ops, { label: "Format" })
  const format = (target: string, patch: StylePatch) =>
    run({ kind: "format", sheet: 0, ranges: [range(target)], patch })
  const style = (ref: string) => {
    const { r0, c0 } = range(ref)
    return wb.styles.get(sheet.styleAt(r0, c0))
  }
  const type = (ref: string, text: string) => {
    const { r0, c0 } = range(ref)
    run(writeInputOp(0, sheet, wb.styles, r0, c0, text))
  }
  return { wb, sheet, run, format, style, type }
}

describe("formatting", () => {
  it("formats a range cell by cell, adding to what each cell has", () => {
    const { format, style } = setup()
    format("A1:B2", { bold: true })
    format("B2:C3", { italic: true })
    expect(style("A1")).toEqual({ bold: true })
    expect(style("B2")).toEqual({ bold: true, italic: true })
    expect(style("C3")).toEqual({ italic: true })
    expect(style("D4")).toEqual({})
  })

  it("a whole column is one entry; its empty cells show it and cells typed later keep it", () => {
    const { sheet, format, style, type } = setup()
    type("B5", "x")
    format("B:B", { fill: "#ffff00" })
    expect(sheet.colStyles.size).toBe(1)
    expect(style("B5")).toEqual({ fill: "#ffff00" }) // the stored cell was formatted too
    expect(style("B900000")).toEqual({ fill: "#ffff00" })
    type("B7", "10%")
    // Typing a percentage adds Percent to the column's format.
    expect(style("B7")).toEqual({ fill: "#ffff00", numberFormat: "0%" })
  })

  it("where a formatted row meets a formatted column the cell takes both, row first (as Excel)", () => {
    const { format, style } = setup()
    format("3:3", { italic: true })
    format("C:C", { bold: true })
    expect(style("C3")).toEqual({ italic: true, bold: true })
    expect(style("C4")).toEqual({ bold: true })
    expect(style("D3")).toEqual({ italic: true })
  })

  it("Clear Formats inside a formatted column keeps the cell plain", () => {
    const { sheet, format, style } = setup()
    format("A:A", { bold: true })
    format("A2", CLEAR_ALL)
    expect(sheet.cells.style(1, 0)).toBe(EXPLICIT_DEFAULT_ID)
    expect(style("A2")).toEqual({})
    expect(style("A3")).toEqual({ bold: true })
  })

  it("undo and redo restore formats exactly", () => {
    const { wb, format, style } = setup()
    format("A1:C3", { bold: true })
    format("B:B", { fill: "#ff0000" })
    format("2:2", CLEAR_ALL)
    expect(style("B2")).toEqual({})
    wb.undo()
    expect(style("B2")).toEqual({ bold: true, fill: "#ff0000" })
    wb.undo()
    expect(style("B2")).toEqual({ bold: true })
    expect(style("B9")).toEqual({})
    wb.redo()
    wb.redo()
    expect(style("B2")).toEqual({})
    expect(style("B9")).toEqual({ fill: "#ff0000" })
  })

  it("inserted rows and columns take the formats of the line before (Format Same As Above)", () => {
    const { wb, run, format, style } = setup()
    format("A1:C1", { bold: true })
    format("2:2", { italic: true })
    run({ kind: "insertRows", sheet: 0, at: 1, count: 2 }) // above row 2: rows 2–3 copy row 1's cells
    expect(style("B2")).toEqual({ bold: true })
    expect(style("B3")).toEqual({ bold: true })
    expect(style("D2")).toEqual({})
    expect(style("A4")).toEqual({ italic: true }) // the formatted row moved down
    run({ kind: "insertRows", sheet: 0, at: 4, count: 1 }) // below row 4: the row format repeats
    expect(style("Z5")).toEqual({ italic: true })
    wb.undo()
    wb.undo()
    expect(style("B2")).toEqual({ italic: true })
    expect(style("A3")).toEqual({})
  })

  it("deleting rows moves row formats and merges, shrinking a merge that loses rows", () => {
    const { wb, sheet, run, format, style } = setup()
    format("10:10", { bold: true })
    run({ kind: "merge", sheet: 0, ranges: [range("B4:C7")] })
    run({ kind: "merge", sheet: 0, ranges: [range("E2:E3")] })
    run({ kind: "deleteRows", sheet: 0, at: 2, count: 3 }) // rows 3–5
    expect(style("A7")).toEqual({ bold: true })
    expect(sheet.merges.all).toEqual([range("B3:C4")])
    wb.undo()
    expect(sheet.merges.all).toContainEqual(range("E2:E3"))
    expect(sheet.merges.all).toContainEqual(range("B4:C7"))
    expect(style("A10")).toEqual({ bold: true })
  })

  it("merging over merged cells absorbs them; unmerge takes every merge it touches", () => {
    const { wb, sheet, run } = setup()
    run({ kind: "merge", sheet: 0, ranges: [range("A1:B2")] })
    run({ kind: "merge", sheet: 0, ranges: [range("A1:D4")] })
    expect(sheet.merges.all).toEqual([range("A1:D4")])
    expect(sheet.merges.at(3, 3)).toEqual(range("A1:D4"))
    run({ kind: "unmerge", sheet: 0, ranges: [range("C3")] })
    expect(sheet.merges.size).toBe(0)
    wb.undo()
    wb.undo()
    expect(sheet.merges.all).toEqual([range("A1:B2")])
  })

  it("format changes do not recalculate", () => {
    const { wb, format, type } = setup()
    type("A1", "1")
    type("A2", "=A1+1")
    const before = wb.calc.formulaCount
    format("A:A", { bold: true })
    format("1:1", { italic: true })
    expect(wb.calc.formulaCount).toBe(before)
    expect(wb.sheets[0].cells.value(1, 0)).toBe(2)
  })
})

describe("merged cells and navigation", () => {
  it("arrows step over a merged cell as one cell and land on its top-left", async () => {
    const { step, growToMerges, isSingleCellOn, moveWithinSelection } = await import("../src/navigation")
    const { selectRange } = await import("../src/selection")
    const { sheet, run } = setup()
    run({ kind: "merge", sheet: 0, ranges: [range("B2:D4")] })
    expect(step(sheet, { row: 2, col: 0 }, "right")).toEqual({ row: 1, col: 1 }) // A3 → the merge (B2)
    expect(step(sheet, { row: 1, col: 1 }, "right")).toEqual({ row: 1, col: 4 }) // out at E2
    expect(step(sheet, { row: 1, col: 1 }, "down")).toEqual({ row: 4, col: 1 }) // out at B5
    expect(growToMerges(sheet, range("A3:C3"))).toEqual(range("A2:D4"))
    expect(isSingleCellOn(sheet, selectRange(range("B2:D4")))).toBe(true)
    // Enter inside a selection skips the merge's other cells.
    const sel = selectRange(range("B2:E4"))
    const next = moveWithinSelection(sheet, sel, "right")
    expect(next.active).toEqual({ row: 1, col: 4 })
  })
})
