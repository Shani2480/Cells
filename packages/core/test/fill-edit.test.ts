import { describe, expect, it } from "vitest"
import { parseRange } from "../src/address"
import type { CellData } from "../src/cell-store"
import { blockFromText, clearContentsOp, fillOp, pasteArea, pasteBlockOp, readBlock } from "../src/edit"
import { fillSeries } from "../src/fill"
import { selectionStats } from "../src/stats"
import { book } from "./helpers"

const v = (...values: (string | number)[]): CellData[] => values.map((value) => ({ value }))
const out = (cells: (CellData | undefined)[]) => cells.map((cell) => cell?.value)

describe("fill series", () => {
  // Every row recorded from Excel's AutoFill (tools/excel-run.py, Excel 16.0 build 17932).
  const excel: [string, (string | number)[], number, (string | number)[]][] = [
    ["one number copies", [5], 3, [5, 5, 5]],
    ["two numbers count", [1, 3], 3, [5, 7, 9]],
    ["three numbers follow their trend", [1, 2, 4], 2, [5.33333333333333, 6.83333333333333]],
    ["text counts its last number", ["Item 1"], 3, ["Item 2", "Item 3", "Item 4"]],
    ["text keeps an even step", ["Item 1", "Item 3"], 2, ["Item 5", "Item 7"]],
    ["text with uneven steps copies", ["Item 1", "Item 2", "Item 4"], 3, ["Item 1", "Item 2", "Item 4"]],
    ["leading zeros stay", ["Q01"], 2, ["Q02", "Q03"]],
    ["quarters wrap", ["Q1"], 5, ["Q2", "Q3", "Q4", "Q1", "Q2"]],
    ["Qtr wraps", ["Qtr1"], 4, ["Qtr2", "Qtr3", "Qtr4", "Qtr1"]],
    ["Quarter wraps", ["Quarter 3"], 3, ["Quarter 4", "Quarter 1", "Quarter 2"]],
    ["Q5 is not a quarter", ["Q5"], 2, ["Q6", "Q7"]],
    ["ordinals", ["1st"], 2, ["2nd", "3rd"]],
    ["ordinals past 20", ["21st"], 3, ["22nd", "23rd", "24th"]],
    ["teens are th", ["11th"], 3, ["12th", "13th", "14th"]],
    ["a spaced suffix is just text", ["1 st"], 2, ["2 st", "3 st"]],
    ["the first number wins", ["2 x 3"], 2, ["3 x 3", "4 x 3"]],
    ["a number inside text copies", ["a1b"], 2, ["a1b", "a1b"]],
    ["days", ["Fri"], 3, ["Sat", "Sun", "Mon"]],
    ["days in capitals", ["MONDAY"], 1, ["TUESDAY"]],
    ["months by two", ["Jan", "Mar"], 2, ["May", "Jul"]],
    ["other text repeats", ["a", "b"], 3, ["a", "b", "a"]],
    ["a falling text series turns at zero", ["Item 3", "Item 1"], 3, ["Item 1", "Item 3", "Item 5"]],
    ["mixed: a number among text counts", ["x", 1], 4, ["x", 2, "x", 3]],
    ["mixed: runs continue separately", ["x", 1, 2], 6, ["x", 3, 4, "x", 5, 6]],
    ["mixed: a run's own trend", ["x", 1, 3], 6, ["x", 5, 7, "x", 9, 11]],
    ["mixed: number first", [1, "x"], 4, [2, "x", 3, "x"]],
    ["mixed: text series and number", ["Item 1", 5], 4, ["Item 2", 6, "Item 3", 7]],
    ["mixed: three runs", [1, 2, "x", 10], 8, [3, 4, "x", 11, 5, 6, "x", 12]],
    ["mixed: day and number", ["Mon", 1], 4, ["Tue", 2, "Wed", 3]],
    ["mixed: plain text and numbered text", ["x", "x2"], 4, ["x", "x3", "x", "x4"]],
    ["TRUE copies", ["TRUE"], 2, ["TRUE", "TRUE"]],
  ]

  it.each(excel)("%s", (_, source, count, expected) => {
    const cells = source.map((value) => ({ value: value === "TRUE" ? true : value }))
    const got = out(fillSeries(cells, count, 1)).map((value) => (value === true ? "TRUE" : value))
    expect(got).toEqual(expected)
  })

  it("fills upward (Excel: 'Item 1' up gives Item 0, Item 1, Item 2)", () => {
    expect(out(fillSeries(v("Item 1"), 3, -1))).toEqual(["Item 0", "Item 1", "Item 2"])
    expect(out(fillSeries(v("Item 0"), 2, -1))).toEqual(["Item 1", "Item 2"])
    expect(out(fillSeries(v("a", "b"), 3, -1))).toEqual(["b", "a", "b"])
    expect(out(fillSeries(v(10, 20), 2, -1))).toEqual([0, -10])
    expect(out(fillSeries(v("Jan"), 1, -1))).toEqual(["Dec"])
  })

  it("Ctrl swaps copying and counting", () => {
    expect(out(fillSeries(v(5), 3, 1, true))).toEqual([6, 7, 8])
    expect(out(fillSeries(v(5), 2, -1, true))).toEqual([4, 3])
    expect(out(fillSeries(v(1, 3), 3, 1, true))).toEqual([1, 3, 1])
    expect(out(fillSeries(v("Item 1"), 2, 1, true))).toEqual(["Item 1", "Item 1"])
  })

  it("repeats formats with the pattern", () => {
    expect(fillSeries([{ value: "x", style: 2 }, undefined], 2, 1)).toEqual([{ value: "x", style: 2 }, undefined])
  })
})

describe("editing commands", () => {
  it("fills a range down, column by column", () => {
    const { wb, sheet } = book({ A1: 1, A2: 2, B1: "Mon", B2: "x" })
    const op = fillOp(0, sheet, parseRange("A1:B2")!, parseRange("A1:B4")!)!
    wb.transact([op], { label: "AutoFill" })
    expect([3, 4].map((r) => sheet.cells.value(r - 1, 0))).toEqual([3, 4])
    expect([3, 4].map((r) => sheet.cells.value(r - 1, 1))).toEqual(["Tue", "x"])
  })

  it("tiles a paste over a multiple of its size", () => {
    const block = blockFromText(book().wb.styles, [["1", "a"]])
    expect(pasteArea(parseRange("C3:F4")!, block)).toEqual(parseRange("C3:F4"))
    expect(pasteArea(parseRange("C3:E4")!, block)).toEqual(parseRange("C3:D3"))
    expect(pasteArea(parseRange("C3")!, block)).toEqual(parseRange("C3:D3"))
    const { wb, sheet } = book()
    wb.transact([pasteBlockOp(0, parseRange("A1:D2")!, block)], { label: "Paste" })
    expect(readBlock(sheet, parseRange("A1:D2")!).cells.map((row) => row.map((c) => c?.value))).toEqual([
      [1, "a", 1, "a"],
      [1, "a", 1, "a"],
    ])
  })

  it("Delete keeps formats", () => {
    const { wb, sheet } = book({ A1: "x" })
    sheet.cells.set(0, 0, { value: "x", style: 4 })
    wb.transact([clearContentsOp(0, sheet, [parseRange("A:A")!])!], { label: "Clear" })
    expect(sheet.cells.get(0, 0)).toEqual({ style: 4 })
    expect(clearContentsOp(0, sheet, [parseRange("A:A")!])).toBeNull()
  })

  it("parses pasted text and sets entry formats", () => {
    const { wb } = book()
    const block = blockFromText(wb.styles, [["10%", "'007", "TRUE"]])
    expect(block.cells[0][0]!.value).toBe(0.1)
    expect(wb.styles.get(block.cells[0][0]!.style!).numberFormat).toBe("0%")
    expect(block.cells[0][1]).toEqual({ value: "007", input: "'007" })
    expect(block.cells[0][2]).toEqual({ value: true })
  })
})

describe("status bar figures", () => {
  it("counts each cell once and sums numbers only", () => {
    const { sheet } = book({ A1: 1, A2: 2, A3: "x", B1: 10 })
    const stats = selectionStats(sheet, [parseRange("A1:B3")!, parseRange("A1:A2")!])
    expect(stats).toMatchObject({ count: 4, numericCount: 3, sum: 13, min: 1, max: 10 })
    expect(stats.average).toBeCloseTo(13 / 3)
    expect(selectionStats(sheet, [parseRange("C1:C9")!]).average).toBeNull()
  })
})
