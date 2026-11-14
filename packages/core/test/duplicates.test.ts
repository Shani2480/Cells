// Remove Duplicates checked against Excel 365 (Range.RemoveDuplicates on the same columns): which rows stayed.
import { describe, expect, it } from "vitest"
import { applyOp, parseInput, removeDuplicatesOps, Workbook } from "../src"

function run(inputs: string[], formats: (string | null)[] = [], values: Record<number, unknown> = {}) {
  const wb = new Workbook()
  const s = wb.sheets[0]
  s.cells.set(0, 0, { value: "H" })
  inputs.forEach((text, i) => {
    const style = formats[i] ? wb.styles.intern({ numberFormat: formats[i]! }) : 0
    const parsed = text === "" ? null : parseInput(text)
    const value = i in values ? values[i] : parsed?.value
    if (value !== undefined)
      s.cells.set(i + 1, 0, {
        value: value as never,
        ...(parsed?.input ? { input: parsed.input } : {}),
        ...(style ? { style } : {}),
      })
    s.cells.set(i + 1, 1, { value: `r${i}` })
  })
  const range = { r0: 0, c0: 0, r1: inputs.length, c1: 1 }
  const result = removeDuplicatesOps(0, s, wb.styles, range, [0], true)
  for (const op of result.ops) applyOp(wb.sheets, op, wb.styles)
  return { kept: inputs.map((_, i) => s.cells.get(i + 1, 1)?.value ?? ""), ...result }
}

describe("Remove Duplicates, as Excel", () => {
  it("text ignores case", () => expect(run(["apple", "Apple", "APPLE"]).kept).toEqual(["r0", "", ""]))
  it("a number isn't the same text", () => expect(run(["1", "'1", "1.0"]).kept).toEqual(["r0", "r1", ""]))
  it("numbers compare exactly", () =>
    expect(run(["0.3", "=0.1+0.2", "0.30000000000000004"], [], { 1: 0.1 + 0.2 }).kept).toEqual(["r0", "r1", ""]))
  it("the same number shown differently isn't a duplicate", () =>
    expect(run(["46100", "46100", "46100"], ["m/d/yyyy", null, "0.00"]).kept).toEqual(["r0", "r1", "r2"]))
  it("nor different numbers shown the same", () =>
    expect(run(["1.001", "1.002", "1"], ["0.0", "0.0", "0.0"]).kept).toEqual(["r0", "r1", "r2"]))
  it("trailing spaces count", () => expect(run(["a", "a ", "a"]).kept).toEqual(["r0", "r1", ""]))
  it("blanks match", () => expect(run(["", "x", "", "x"]).kept).toEqual(["r0", "r1", "", ""]))
  it("TRUE and the text TRUE match", () =>
    expect(run(["TRUE", "'TRUE", "=1=1"], [], { 2: true }).kept).toEqual(["r0", "", ""]))
  it("errors match by code, not the text #N/A", () =>
    expect(run(["=NA()", "#N/A", "'#N/A"], [], { 0: parseInput("#N/A")!.value }).kept).toEqual(["r0", "r2", ""]))

  it("several columns; cells beside and below the range stay", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    const rows = [
      ["K", "V", "below"],
      ["a", "1", "c0"],
      ["b", "2", "c1"],
      ["a", "1", "c2"],
      ["b", "3", "c3"],
      ["a", "1", "c4"],
    ]
    rows.forEach((r, i) => r.forEach((v, c) => s.cells.set(i, c, { value: parseInput(v)!.value })))
    s.cells.set(7, 0, { value: "under" })
    const result = removeDuplicatesOps(0, s, wb.styles, { r0: 0, c0: 0, r1: 5, c1: 1 }, [0, 1], true)
    expect([result.removed, result.remain]).toEqual([2, 3])
    for (const op of result.ops) applyOp(wb.sheets, op, wb.styles)
    const row = (i: number) => [0, 1, 2].map((c) => s.cells.get(i, c)?.value ?? "").join("/")
    expect([1, 2, 3, 4, 5, 6, 7].map(row)).toEqual(["a/1/c0", "b/2/c1", "b/3/c2", "//c3", "//c4", "//", "under//"])
  })
})
