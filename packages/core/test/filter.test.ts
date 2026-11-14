// AutoFilter checked against Excel 365: fixtures/excel-filter.json lists, for each filter, the rows (0-based, under
// the header) Excel left visible on the same column of mixed values.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  errorValue,
  filterRowEntries,
  filterVisibility,
  wildcardRegExp,
  Workbook,
  type AutoFilter,
  type ColumnFilter,
  type FilterCondition,
  type FilterOperator,
} from "../src"

const fx = JSON.parse(readFileSync(new URL("./fixtures/excel-filter.json", import.meta.url), "utf8"))

/** The probe's column: what Excel held in A2:A23. */
function probe(withError = true) {
  const wb = new Workbook()
  const s = wb.sheets[0]
  const text = wb.styles.intern({ numberFormat: "@" })
  const date = wb.styles.intern({ numberFormat: "m/d/yyyy" })
  s.cells.set(0, 0, { value: "Val" })
  const values: unknown[] = [
    5,
    10,
    3.5,
    "apple",
    "banana",
    "Anna",
    "t5",
    true,
    "#NA",
    "",
    "date",
    -2,
    "b",
    "B",
    100,
    0,
    "pan",
    "x*y",
    "a?c",
    "abc",
    5.0000000001,
    "  apple",
  ]
  values.forEach((v, i) => {
    const r = i + 1
    if (typeof v === "number" || typeof v === "boolean") s.cells.set(r, 0, { value: v })
    else if (v === "#NA") {
      if (withError) s.cells.set(r, 0, { value: errorValue("#N/A"), input: "=NA()" })
    } else if (v === "date") s.cells.set(r, 0, { value: 46095, style: date })
    else if (v === "t5") s.cells.set(r, 0, { value: "5", style: text })
    else if (v !== "") s.cells.set(r, 0, { value: v as string, style: text })
  })
  return { wb, s, range: { r0: 0, c0: 0, r1: values.length, c1: 0 } }
}

const OPS: [string, FilterOperator][] = [
  ["<>", "ne"],
  [">=", "ge"],
  ["<=", "le"],
  [">", "gt"],
  ["<", "lt"],
  ["=", "eq"],
]
/** Excel's criterion text (">=5") as a condition. */
function condition(text: string): FilterCondition {
  const [prefix, op] = OPS.find(([p]) => text.startsWith(p))!
  return { op, value: text.slice(prefix.length) }
}

function visible(column: ColumnFilter, withError = true) {
  const { wb, s, range } = probe(withError)
  const filter: AutoFilter = { range, columns: { 0: column } }
  return filterVisibility(s, wb.styles, filter, 46095)
    .map((shown, i) => (shown ? i : -1))
    .filter((i) => i >= 0)
}

describe("AutoFilter, as Excel's", () => {
  it.each(Object.entries(fx.custom as Record<string, number[]>))("custom %s", (criterion, rows) => {
    expect(visible({ kind: "custom", conditions: [condition(criterion)] })).toEqual(rows)
  })

  it("two conditions with And and Or", () => {
    expect(visible({ kind: "custom", conditions: fx.and.conditions.map(condition), and: true })).toEqual(fx.and.rows)
    expect(visible({ kind: "custom", conditions: fx.or.conditions.map(condition), and: false })).toEqual(fx.or.rows)
  })

  it("value lists match the shown text", () => {
    for (const v of fx.values) expect(visible({ kind: "values", values: v.values, blanks: v.blanks })).toEqual(v.rows)
  })

  it("Top 10 by items and percent, ties kept (Excel's results without the error cell)", () => {
    const w = fx.withoutError
    expect(visible({ kind: "top10", top: true, percent: false, count: 3 }, false)).toEqual(w.top3)
    expect(visible({ kind: "top10", top: false, percent: false, count: 2 }, false)).toEqual(w.bottom2)
    expect(visible({ kind: "top10", top: true, percent: true, count: 25 }, false)).toEqual(w.top25pct)
    expect(visible({ kind: "top10", top: false, percent: true, count: 40 }, false)).toEqual(w.bottom40pct)
  })

  it("above and below average look at numbers only", () => {
    // Numbers: 5, 10, 3.5, 46095 (a date), -2, 100, 0, 5.0000000001: the average is 5777.4375.
    expect(visible({ kind: "dynamic", type: "aboveAverage" })).toEqual([10])
    expect(visible({ kind: "dynamic", type: "belowAverage" })).toEqual([0, 1, 2, 11, 14, 15, 20])
  })

  it("date filters relative to today, and dates checked by year and month", () => {
    const { wb, s } = probe()
    const date = wb.styles.intern({ numberFormat: "m/d/yyyy" })
    const dates = [46094, 46095, 46096, 46113, 46022, 45700]
    dates.forEach((d, i) => s.cells.set(i + 1, 1, { value: d, style: date }))
    const range = { r0: 0, c0: 1, r1: dates.length, c1: 1 }
    const rows = (column: ColumnFilter) =>
      filterVisibility(s, wb.styles, { range, columns: { 0: column } }, 46095)
        .map((v, i) => (v ? i : -1))
        .filter((i) => i >= 0)
    // Today is 3/14/2026, a Saturday.
    expect(rows({ kind: "dynamic", type: "today" })).toEqual([1])
    expect(rows({ kind: "dynamic", type: "yesterday" })).toEqual([0])
    expect(rows({ kind: "dynamic", type: "thisWeek" })).toEqual([0, 1])
    expect(rows({ kind: "dynamic", type: "nextWeek" })).toEqual([2])
    expect(rows({ kind: "dynamic", type: "thisMonth" })).toEqual([0, 1, 2])
    expect(rows({ kind: "dynamic", type: "nextMonth" })).toEqual([3])
    expect(rows({ kind: "dynamic", type: "lastQuarter" })).toEqual([4])
    expect(rows({ kind: "dynamic", type: "lastYear" })).toEqual([4, 5])
    expect(rows({ kind: "dynamic", type: "M4" })).toEqual([3])
    expect(rows({ kind: "values", values: [], dates: [{ year: 2026, month: 3 }] })).toEqual([0, 1, 2])
  })

  it("by colour, and only the rows whose state changes are written", () => {
    const { wb, s, range } = probe()
    const red = wb.styles.intern({ fill: "#FF0000" })
    s.cells.set(1, 0, { value: 5, style: red })
    const filter: AutoFilter = { range, columns: { 0: { kind: "color", cell: true, color: "#ff0000" } } }
    const entries = filterRowEntries(s, wb.styles, filter, 46095)
    expect(entries.filter((e) => !e.hidden)).toEqual([])
    expect(entries).toHaveLength(21)
    expect(entries.find((e) => e.index === 1)).toBeUndefined()
  })

  it("wildcards", () => {
    expect(wildcardRegExp("*an*").test("banana")).toBe(true)
    expect(wildcardRegExp("x~*y").test("x*y")).toBe(true)
    expect(wildcardRegExp("x~*y").test("xay")).toBe(false)
    expect(wildcardRegExp("a.c").test("abc")).toBe(false)
  })
})
