// Find and Replace checked against Excel 365: fixtures/excel-find.json holds what Range.Find / FindNext found (in
// order, starting after C4) and what Range.Replace left, on the probe sheet built below.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  cellsToRanges,
  editTextFor,
  findAll,
  findNext,
  parseInput,
  replacedText,
  searchText,
  specialCells,
  Workbook,
  type FindOptions,
} from "../src"

const fx = JSON.parse(readFileSync(new URL("./fixtures/excel-find.json", import.meta.url), "utf8"))
const name = (c: { row: number; col: number }) => String.fromCharCode(65 + c.col) + (c.row + 1)

function probe() {
  const wb = new Workbook()
  const s = wb.sheets[0]
  const put = (a: string, input: string, numberFormat?: string) => {
    const col = a.charCodeAt(0) - 65
    const row = Number(a.slice(1)) - 1
    const parsed = parseInput(input)!
    const style = numberFormat ? wb.styles.intern({ numberFormat }) : 0
    s.cells.set(row, col, {
      value: parsed.value,
      ...(parsed.input ? { input: parsed.input } : {}),
      ...(style ? { style } : {}),
    })
  }
  put("A1", "Apple pie")
  put("B1", "apple")
  put("C1", "1234.5", "$#,##0.00")
  put("A2", "=SUM(C1,10)")
  put("B2", "10")
  put("C2", "pineapple")
  put("A3", "x*y")
  put("B3", "100", "0%")
  put("C3", '=B1&"!"')
  put("A4", "TRUE")
  put("B4", "a?c")
  put("C4", "46100", "d-mmm-yy")
  // Values Excel calculated.
  s.cells.set(1, 0, { value: 1244.5, input: "=SUM(C1,10)" })
  s.cells.set(2, 2, { value: "apple!", input: '=B1&"!"' })
  s.annotations.set(1, 1, { kind: "note", author: "", text: "apple note", shown: false })
  return { wb, s }
}

describe("Find, as Excel finds", () => {
  const { wb, s } = probe()
  const area = [{ r0: 0, c0: 0, r1: 3, c1: 2 }]
  for (const q of fx.find as (FindOptions & { cells: string[] })[]) {
    it(`${q.text} in ${q.lookIn}${q.entire ? ", whole cell" : ""}${q.matchCase ? ", case" : ""}${q.byColumns ? ", by columns" : ""}`, () => {
      expect(findAll(s, wb.styles, q, area).map(name)).toEqual(q.cells)
    })
  }

  it("Find Next goes round from the active cell; Shift goes back", () => {
    const o: FindOptions = { text: "apple", lookIn: "formulas" }
    expect(findNext(s, wb.styles, o, area, { row: 0, col: 1 })).toEqual({ row: 1, col: 2 })
    expect(findNext(s, wb.styles, o, area, { row: 1, col: 2 })).toEqual({ row: 0, col: 0 })
    expect(findNext(s, wb.styles, o, area, { row: 0, col: 0 }, true)).toEqual({ row: 1, col: 2 })
  })
})

describe("Replace, as Excel replaces", () => {
  const { wb, s } = probe()
  for (const r of fx.replace) {
    it(`${r.text} → ${r.with}${r.entire ? ", whole cell" : ""}${r.matchCase ? ", case" : ""}`, () => {
      const o: FindOptions = { text: r.text, lookIn: "formulas", entire: r.entire, matchCase: r.matchCase }
      for (const [a, want] of Object.entries(r.cells as Record<string, { input: string; type: string }>)) {
        const col = a.charCodeAt(0) - 69
        const row = Number(a.slice(1)) - 1
        const before = searchText(s, wb.styles, row, col, "formulas")!
        const after = replacedText(before, o, r.with) ?? before
        // Excel's copy sat four columns to the right: its references moved with it.
        const expected = want.input.replace(
          /\b([E-G])(\d)/g,
          (_m, c: string, d: string) => String.fromCharCode(c.charCodeAt(0) - 4) + d,
        )
        if (want.type === "Double" && !expected.startsWith("="))
          expect(parseInput(after)?.value, a).toBe(Number(expected))
        else if (want.type === "Boolean") expect(after, a).toBe(expected)
        else expect(after, a).toBe(expected)
      }
    })
  }
})

describe("the formula bar's text", () => {
  it("shows dates, times and percentages as Excel's formula bar does", () => {
    expect(editTextFor(46100, "d-mmm-yy")).toBe("3/19/2026")
    expect(editTextFor(46100.5, "m/d/yyyy")).toBe("3/19/2026 12:00:00 PM")
    expect(editTextFor(0.75, "h:mm")).toBe("6:00:00 PM")
    expect(editTextFor(1, "0%")).toBe("100%")
    expect(editTextFor(0.125, "0.0%")).toBe("12.5%")
    expect(editTextFor(1234.5, "$#,##0.00")).toBe("1234.5")
  })
})

describe("Go To Special", () => {
  const { wb, s } = probe()
  const one = [{ r0: 0, c0: 0, r1: 0, c1: 0 }]
  const at = { row: 0, col: 0 }
  const kinds = { numbers: true, text: true, logicals: true, errors: true }
  it("constants and formulas by type", () => {
    expect(specialCells(s, wb.styles, { kind: "formulas", types: kinds }, one, at)).toEqual([
      { r0: 1, c0: 0, r1: 1, c1: 0 },
      { r0: 2, c0: 2, r1: 2, c1: 2 },
    ])
    // Numbers and logicals only: C1, B2, B3, A4, C4 (the date is a number).
    expect(specialCells(s, wb.styles, { kind: "constants", types: { ...kinds, text: false } }, one, at)).toEqual([
      { r0: 0, c0: 2, r1: 0, c1: 2 },
      { r0: 1, c0: 1, r1: 2, c1: 1 },
      { r0: 3, c0: 0, r1: 3, c1: 0 },
      { r0: 3, c0: 2, r1: 3, c1: 2 },
    ])
  })

  it("precedents and dependents", () => {
    expect(specialCells(s, wb.styles, { kind: "precedents", all: false }, one, { row: 1, col: 0 })).toEqual([
      { r0: 0, c0: 2, r1: 0, c1: 2 },
    ])
    expect(specialCells(s, wb.styles, { kind: "dependents", all: false }, one, { row: 0, col: 1 })).toEqual([
      { r0: 2, c0: 2, r1: 2, c1: 2 },
    ])
  })

  it("blanks, notes, row differences and the last cell", () => {
    s.cells.set(5, 3, { value: 1 })
    expect(specialCells(s, wb.styles, { kind: "lastCell" }, one, at)).toEqual([{ r0: 5, c0: 3, r1: 5, c1: 3 }])
    expect(specialCells(s, wb.styles, { kind: "blanks" }, [{ r0: 3, c0: 2, r1: 5, c1: 3 }], at)).toEqual([
      { r0: 3, c0: 3, r1: 3, c1: 3 },
      { r0: 4, c0: 2, r1: 4, c1: 3 },
      { r0: 5, c0: 2, r1: 5, c1: 2 },
    ])
    expect(specialCells(s, wb.styles, { kind: "notes" }, one, at)).toEqual([{ r0: 1, c0: 1, r1: 1, c1: 1 }])
  })
})

describe("cellsToRanges", () => {
  it("joins cells into rectangles", () => {
    const cells = [0, 1, 2].flatMap((row) => [3, 4].map((col) => ({ row, col })))
    expect(cellsToRanges([...cells, { row: 5, col: 0 }])).toEqual([
      { r0: 0, c0: 3, r1: 2, c1: 4 },
      { r0: 5, c0: 0, r1: 5, c1: 0 },
    ])
  })
})
