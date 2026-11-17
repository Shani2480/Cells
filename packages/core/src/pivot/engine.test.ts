import { describe, expect, it } from "vitest"
import { cellName } from "../address"
import { dateSerial } from "../formula/dates"
import { errorValue } from "../values"
import excel from "./excel-layouts.json"
import { pivotLayout, type PivotAxisItem } from "./engine"
import {
  newPivotField,
  VALUES_FIELD as V,
  type PivotCache,
  type PivotDataField,
  type PivotField,
  type PivotTable,
  type PivotValue,
} from "./model"

/**
 * The engine against Excel 365: the same made-up sales data and the same PivotTables as
 * packages/io/test/fixtures/excel/pivot-basic.xlsx and pivot-more.xlsx (made by tools/excel-run.py), whose cells,
 * locations and row and column items Excel wrote are in excel-layouts.json.
 */

type Expected = {
  cells: Record<string, string | number | { error: string }>
  location: Record<string, string>
  rowItems: string[]
  colItems: string[]
}
const EXCEL = excel as Record<string, Expected>

const REGIONS = ["North", "South", "East", "West"]
const PRODUCTS = ["Apples", "Pears", "Plums"]
const CHANNELS = ["Shop", "Online"]
const NAMES = ["Region", "Product", "Date", "Units", "Amount", "Channel"]

/** The Data sheet of the reference files (pivot-more changes two cells). */
function cache(more: boolean): PivotCache {
  const records: PivotValue[][] = []
  for (let i = 0; i < 24; i++) {
    records.push([
      REGIONS[(i * 7) % 4],
      PRODUCTS[(i * 5) % 3],
      dateSerial(2025, (i % 12) + 1, ((i * 3) % 27) + 1),
      ((i * 13) % 17) + 1,
      i === 9 ? null : ((i * 37) % 53) * 12.5 - 40,
      CHANNELS[(i * 3) % 2],
    ])
  }
  if (more) {
    records[5][4] = errorValue("#N/A")
    records[6][1] = "Pears"
  }
  // Shared items in the order first met, as Excel keeps them.
  const fields = NAMES.map((name, f) => {
    const items: PivotValue[] = []
    for (const r of records) if (!items.includes(r[f])) items.push(r[f])
    return { name, items }
  })
  return { id: 1, source: { kind: "range", sheet: "Data", ref: "A1:F25" }, fields, records, refreshOnLoad: false }
}

type Layout = "compact" | "outline" | "tabular"

/** A PivotTable as Excel makes it: fields' items in alphabetical order, then the settings each sheet changed. */
function table(
  c: PivotCache,
  o: {
    at: string
    rows?: number[]
    cols?: number[]
    pages?: { field: number; item: number | null }[]
    values: PivotDataField[]
    layout?: Layout
    field?: Record<number, Partial<PivotField>>
    table?: Partial<PivotTable>
  },
): PivotTable {
  const layout = o.layout ?? "compact"
  const compact = layout === "compact"
  const outline = layout !== "tabular"
  const fields = c.fields.map((cf, f) => {
    const field = newPivotField(cf.items!.length, { compact, outline })
    const sorted = cf.items!.map((v, x) => ({ v, x })).sort((a, b) => String(a.v).localeCompare(String(b.v)))
    field.items = sorted.map(({ x }) => ({ x }))
    if (typeof cf.items![0] === "number")
      field.items = cf
        .items!.map((v, x) => ({ v: v as number, x }))
        .sort((a, b) => a.v - b.v)
        .map(({ x }) => ({ x }))
    // Excel keeps a sorted field's items in the order shown.
    if (o.field?.[f]?.sort === "descending") field.items.reverse()
    return { ...field, ...o.field?.[f] }
  })
  const m = /^([A-Z]+)(\d+)$/.exec(o.at)!
  return {
    name: "PT",
    cacheId: 1,
    at: { row: Number(m[2]) - 1, col: m[1].charCodeAt(0) - 65 },
    fields,
    rows: o.rows ?? [],
    cols: o.cols ?? [],
    pages: o.pages ?? [],
    values: o.values,
    rowGrandTotals: true,
    colGrandTotals: true,
    compact,
    outline,
    compactData: compact,
    captions: { data: "Values" },
    showError: false,
    showMissing: true,
    style: {
      name: "PivotStyleLight16",
      rowHeaders: true,
      colHeaders: true,
      rowStripes: false,
      colStripes: false,
      lastColumn: true,
    },
    ...o.table,
  }
}

const itemText = (it: PivotAxisItem) => {
  const attrs = [
    it.t !== "data" ? `t="${it.t === "count" ? "countA" : it.t}"` : "",
    it.r ? `r="${it.r}"` : "",
    it.i ? `i="${it.i}"` : "",
  ].filter(Boolean)
  return `${attrs.join(" ")}[${it.x.join(",")}]`
}

/** Compares the engine's layout with what Excel wrote for the sheet. */
function check(sheet: string, t: PivotTable, c: PivotCache) {
  const want = EXCEL[sheet]
  const got = pivotLayout(t, c)
  const ref = `${cellName({ row: got.range.r0, col: got.range.c0 })}:${cellName({ row: got.range.r1, col: got.range.c1 })}`
  expect({
    ref,
    firstHeaderRow: String(got.firstHeaderRow),
    firstDataRow: String(got.firstDataRow),
    firstDataCol: String(got.firstDataCol),
  }).toEqual({
    ref: want.location.ref,
    firstHeaderRow: want.location.firstHeaderRow,
    firstDataRow: want.location.firstDataRow,
    firstDataCol: want.location.firstDataCol,
  })
  expect(got.rowItems.map(itemText)).toEqual(want.rowItems)
  expect(got.colItems.map(itemText)).toEqual(want.colItems)
  const cells: Record<string, unknown> = {}
  for (const cell of got.cells) if (cell.value !== undefined) cells[cellName(cell)] = cell.value
  expect(Object.keys(cells).sort()).toEqual(Object.keys(want.cells).sort())
  for (const [at, value] of Object.entries(want.cells)) {
    if (typeof value === "number") {
      expect(typeof cells[at], at).toBe("number")
      expect(cells[at] as number, at).toBeCloseTo(value, Math.abs(value) > 1e6 ? -6 : 9)
    } else expect(cells[at], at).toEqual(value)
  }
}

const sum = (field: number, name: string): PivotDataField => ({ name, field, fn: "sum" })
const AMOUNT = 4
const UNITS = 3

describe("PivotTable engine against Excel 365", () => {
  const basic = cache(false)
  const more = cache(true)

  it("compact form with subtotals on each item's line", () => {
    check("Compact", table(basic, { at: "A3", rows: [0, 1], values: [sum(AMOUNT, "Sum of Amount")] }), basic)
  })

  it("outline form, subtotals at the bottom, two value fields", () => {
    const t = table(basic, {
      at: "A3",
      layout: "outline",
      rows: [0, 1],
      cols: [V],
      values: [sum(AMOUNT, "Sum of Amount"), { name: "Count of Units", field: UNITS, fn: "count" }],
      field: { 0: { subtotalTop: false } },
    })
    check("Outline", t, basic)
  })

  it("tabular form with repeated labels, no subtotals", () => {
    const t = table(basic, {
      at: "A3",
      layout: "tabular",
      rows: [0, 1],
      values: [{ name: "Average of Units", field: UNITS, fn: "average" }],
      field: Object.fromEntries(
        [0, 1, 2, 3, 4, 5].map((f) => [f, { repeatLabels: true, ...(f === 0 ? { subtotals: [] } : {}) }]),
      ),
      table: { colGrandTotals: false },
    })
    check("Tabular", t, basic)
  })

  it("rows by columns, a report filter, a hidden item and a descending sort", () => {
    const t = table(basic, {
      at: "A5",
      rows: [1],
      cols: [0],
      pages: [{ field: 5, item: 0 }],
      values: [sum(UNITS, "Sum of Units")],
      field: { 0: { sort: "descending" } },
    })
    t.fields[1].items.find((i) => basic.fields[1].items![i.x] === "Plums")!.hidden = true
    check("Cross", t, basic)
  })

  it("every summary function", () => {
    const fns = [
      ["max", "Max"],
      ["min", "Min"],
      ["countNums", "Count Numbers"],
      ["stdDev", "StdDev"],
      ["stdDevp", "StdDevp"],
      ["var", "Var"],
      ["varp", "Varp"],
    ] as const
    const t = table(basic, {
      at: "A3",
      rows: [0],
      cols: [V],
      values: [
        ...fns.map(([fn, name]) => ({ name: `${name} of Amount`, field: AMOUNT, fn })),
        { name: "Product of Units", field: UNITS, fn: "product" as const },
      ],
    })
    check("Functions", t, basic)
  })

  it("value fields on the rows", () => {
    const t = table(basic, {
      at: "A3",
      rows: [0, V],
      cols: [1],
      values: [sum(AMOUNT, "Sum of Amount"), sum(UNITS, "Sum of Units")],
    })
    check("ValuesRows", t, basic)
  })

  it("two report filters, one with several items picked", () => {
    const t = table(more, {
      at: "A6",
      rows: [1],
      pages: [
        { field: 5, item: null },
        { field: 0, item: null },
      ],
      values: [sum(UNITS, "Sum of Units")],
      field: { 0: { multipleItems: true } },
    })
    t.fields[0].items.find((i) => more.fields[0].items![i.x] === "East")!.hidden = true
    check("TwoPages", t, more)
  })

  it("two column fields with value fields inside, and errors in the data", () => {
    const t = table(more, {
      at: "A3",
      rows: [0],
      cols: [5, 1, V],
      values: [sum(UNITS, "Sum of Units"), { name: "Max of Amount", field: AMOUNT, fn: "max" }],
    })
    check("TwoCols", t, more)
  })

  it("numbers as items, sorted descending, with a grand total caption", () => {
    const t = table(more, {
      at: "A3",
      rows: [UNITS],
      values: [sum(AMOUNT, "Sum of Amount")],
      field: { [UNITS]: { sort: "descending" } },
      table: { captions: { data: "Values", grandTotal: "All of it" } },
    })
    check("Numbers", t, more)
  })

  it("custom subtotals, show items with no data, blank lines", () => {
    const t = table(more, {
      at: "A3",
      layout: "outline",
      rows: [5, 1, 0],
      values: [sum(UNITS, "Sum of Units")],
      field: { 5: { subtotals: ["sum", "count"] }, 1: { showAll: true, insertBlankRow: true } },
      table: { captions: { data: "Values", error: "err", missing: "-" }, showError: true },
    })
    check("Subtotals", t, more)
  })

  it("a collapsed item and captions", () => {
    const t = table(more, {
      at: "A3",
      rows: [0, 1],
      values: [{ name: "Average of Amount", field: AMOUNT, fn: "average" }],
      field: { 0: { name: "Area" } },
      table: { captions: { data: "Values", rowHeader: "Where" } },
    })
    t.fields[0].items.find((i) => more.fields[0].items![i.x] === "North")!.collapsed = true
    check("Collapsed", t, more)
  })

  it("only column fields", () => {
    const t = table(more, {
      at: "A3",
      cols: [1, V],
      values: [sum(UNITS, "Sum of Units"), { name: "Count of Units2", field: UNITS, fn: "count" }],
    })
    check("ColOnly", t, more)
  })
})

describe("PivotTable engine speed", () => {
  it("lays out 100,000 records by region, product and channel quickly", () => {
    const c = cache(false)
    const base = c.records
    c.records = Array.from({ length: 100_000 }, (_, i) => {
      const r = [...base[i % base.length]]
      r[3] = i % 97
      return r
    })
    const t = table(c, {
      at: "A3",
      rows: [0, 1],
      cols: [5],
      values: [sum(UNITS, "Sum of Units"), { name: "Average of Amount", field: AMOUNT, fn: "average" }],
    })
    t.cols = [5, V]
    const start = performance.now()
    const layout = pivotLayout(t, c)
    const ms = performance.now() - start
    expect(layout.rowItems.length).toBe(17)
    expect(ms).toBeLessThan(1500)
  })
})
