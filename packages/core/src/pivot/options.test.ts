import { describe, expect, it } from "vitest"
import { rangeName } from "../address"
import { pivotLayout } from "./engine"
import { pivotChoices, withFieldFilter, withSort } from "./filter"
import { newPivotTable, withFieldIn } from "./make"
import type { PivotCache, PivotTable } from "./model"
import { pivotOptions, withPivotOptions, withReportLayout } from "./settings"
import { pivotArea } from "./sheet"
import { Sheet } from "../sheet"

/** The data of Excel's own runs (PivotTable Options measured in Excel 365). */
const cache: PivotCache = {
  id: 1,
  source: { kind: "range", sheet: "Data", ref: "A1:F7" },
  fields: [
    { name: "Product", items: ["Pens", "Ink", "Paper", "Clips"] },
    { name: "Region", items: ["East", "West"] },
    { name: "Amount", items: null },
    { name: "Year", items: [2024, 2025] },
    { name: "Rep", items: ["Ann", "Bo", "Cy"] },
    { name: "Kind", items: ["A", "B"] },
  ],
  records: [
    ["Pens", "East", 10, 2024, "Ann", "A"],
    ["Ink", "West", 40, 2025, "Bo", "B"],
    ["Paper", "East", 25, 2024, "Ann", "A"],
    ["Pens", "West", 5, 2025, "Cy", "B"],
    ["Clips", "East", 1, 2024, "Bo", "A"],
    ["Ink", "East", 2, 2025, "Cy", "B"],
  ],
  refreshOnLoad: false,
}

/** Region and Product on rows, Kind and Year on columns, Sum of Amount, at H12 as in Excel's run. */
function table(cols = [5, 3], values = 1): PivotTable {
  let t = newPivotTable("PT", cache, { row: 11, col: 7 })
  for (const f of [1, 0]) t = withSort(withFieldIn(t, cache, f, "rows"), cache, f, "ascending")
  for (const f of cols) t = withSort(withFieldIn(t, cache, f, "cols"), cache, f, "ascending")
  t = withFieldIn(t, cache, 2, "values")
  if (values > 1) t = withFieldIn(t, cache, 2, "values")
  return values > 1 ? { ...t, values: [t.values[0], { ...t.values[1], fn: "count", name: "Count of Amount" }] } : t
}

const merges = (t: PivotTable) =>
  pivotLayout(t, cache)
    .merges.map(rangeName)
    .sort((a, b) => a.localeCompare(b, "en", { numeric: true }))
const sorted = (list: string[]) => [...list].sort((a, b) => a.localeCompare(b, "en", { numeric: true }))

describe("Merge and center cells with labels", () => {
  it("merges as Excel 365 does in tabular form: outer items down, totals across, column totals down", () => {
    const t = { ...withReportLayout(table(), "tabular"), mergeLabels: true as const }
    expect(merges(t)).toEqual(
      sorted(["K13:K14", "M13:M14", "N13:N14", "H15:H18", "H19:I19", "H20:H21", "H22:I22", "H23:I23"]),
    )
  })

  it("merges the column labels across their items, with two value fields", () => {
    const t = { ...withReportLayout(table([5, 3], 2), "tabular"), mergeLabels: true as const }
    expect(merges(t)).toEqual(
      sorted([
        "J13:K13",
        "L13:L15",
        "M13:M15",
        "N13:O13",
        "P13:P15",
        "Q13:Q15",
        "R13:R15",
        "S13:S15",
        "J14:K14",
        "N14:O14",
        "H16:H19",
        "H20:I20",
        "H21:H22",
        "H23:I23",
        "H24:I24",
      ]),
    )
  })

  it("in compact form only merges the column totals, and nothing while it's off", () => {
    expect(merges({ ...table(), mergeLabels: true })).toEqual(sorted(["J13:J14", "L13:L14", "M13:M14"]))
    expect(merges(withReportLayout(table(), "tabular"))).toEqual([])
  })
})

describe("compact indent and the report filter area", () => {
  it("indents each level by the characters set", () => {
    const t = { ...table([5]), indent: 3 }
    const clips = pivotLayout(t, cache).cells.find((c) => c.value === "Clips")
    expect(clips?.indent).toBe(3)
    expect(pivotLayout({ ...t, indent: 0 }, cache).cells.find((c) => c.value === "Clips")?.indent).toBeUndefined()
  })

  /** Where each report filter's label lands (relative to the table's corner), in `pages` order. */
  const spots = (patch: Partial<PivotTable>) => {
    const t: PivotTable = { ...table([]), pages: [5, 4, 3].map((field) => ({ field, item: null })), ...patch }
    const layout = pivotLayout(t, cache)
    return {
      spots: layout.cells
        .filter((c) => c.role === "pageLabel")
        .map((c) => `${c.value}@${c.row - layout.range.r0},${c.col - layout.range.c0}`),
      area: rangeName(pivotArea({ ...t, location: layout.range })),
    }
  }

  it("lays the filters out down then over, or over then down, wrapping as Excel 365 did", () => {
    // Excel: Kind, Rep, Year top down; wrap 2 puts Year in a second column three cells over.
    expect(spots({}).spots).toEqual(["Kind@-4,0", "Rep@-3,0", "Year@-2,0"])
    expect(spots({ pageWrap: 2 }).spots).toEqual(["Kind@-3,0", "Rep@-2,0", "Year@-3,3"])
    expect(spots({ pageOverThenDown: true })).toEqual({
      spots: ["Kind@-2,0", "Rep@-2,3", "Year@-2,6"],
      area: "H10:O21",
    })
    expect(spots({ pageOverThenDown: true, pageWrap: 2 }).spots).toEqual(["Kind@-3,0", "Rep@-3,3", "Year@-2,0"])
    expect(spots({ pageWrap: 1 }).spots).toEqual(spots({ pageOverThenDown: true }).spots)
  })
})

describe("Totals & Filters and the other options", () => {
  const label = { kind: "label", op: "beginsWith", value: "P" } as const
  const value = { kind: "value", op: "greaterThan", measure: 0, value: 6 } as const

  it("keeps a label filter and a value filter on one field when multiple filters are allowed", () => {
    const one = withFieldFilter(withFieldFilter(table([]), 0, label), 0, value)
    expect(one.filters?.map((f) => f.kind)).toEqual(["value"])
    const t = { ...table([]), multipleFieldFilters: true as const }
    const both = withFieldFilter(withFieldFilter(t, 0, label), 0, value)
    expect(both.filters?.map((f) => f.kind)).toEqual(["label", "value"])
    expect(pivotChoices(both, cache, 0)?.filters).toHaveLength(2)
    // Paper and Pens begin with P; of those, East's Paper (25) and Pens (10) are over 6, West's Pens (5) isn't.
    const rows = pivotLayout(both, cache)
      .cells.filter((c) => c.role === "rowItem")
      .map((c) => c.value)
    expect(rows).toEqual(["East", "Paper", "Pens"])
    // Clearing Value Filters leaves the label filter.
    expect(withFieldFilter(both, 0, null, "value").filters?.map((f) => f.kind)).toEqual(["label"])
  })

  it("sorts days and months first, by Excel's custom lists, unless that is off", () => {
    const k: PivotCache = {
      ...cache,
      fields: [{ name: "Month", items: ["Mar", "Jan", "Feb", "Apple", "Monday"] }, cache.fields[1], cache.fields[2]],
      records: [
        ["Mar", "East", 1],
        ["Jan", "East", 2],
        ["Feb", "West", 3],
        ["Apple", "West", 4],
        ["Monday", "West", 5],
      ],
    }
    let t = newPivotTable("PT", k, { row: 0, col: 0 })
    t = withSort(withFieldIn(withFieldIn(t, k, 0, "rows"), k, 2, "values"), k, 0, "ascending")
    const labels = (x: PivotTable) =>
      pivotLayout(x, k)
        .cells.filter((c) => c.role === "rowItem")
        .map((c) => c.value)
    // Excel 365: the lists' items first, days before months, then the rest A to Z.
    expect(labels(t)).toEqual(["Monday", "Jan", "Feb", "Mar", "Apple"])
    expect(labels({ ...t, customListSort: false })).toEqual(["Apple", "Feb", "Jan", "Mar", "Monday"])
  })

  it("keeps only what differs from Excel's defaults, and Classic layout turns tabular form on", () => {
    const t = table()
    const sheet = new Sheet("S")
    sheet.pivots = [t]
    const o = pivotOptions(t, cache)
    expect(withPivotOptions(t, sheet, o)).toEqual(t)
    const next = withPivotOptions(t, sheet, {
      ...o,
      mergeLabels: true,
      indent: 4,
      pageWrap: 3,
      preserveFormatting: false,
      gridDropZones: true,
      itemPrintTitles: false,
      fieldPrintTitles: true,
      altText: "Sales",
      altTextSummary: "",
    }) as PivotTable
    expect(next).toMatchObject({
      mergeLabels: true,
      indent: 4,
      pageWrap: 3,
      preserveFormatting: false,
      gridDropZones: true,
      itemPrintTitles: false,
      fieldPrintTitles: true,
      altText: "Sales",
    })
    expect(next.altTextSummary).toBeUndefined()
    expect(next.fields.every((f) => !f.compact && !f.outline)).toBe(true)
    expect(pivotOptions(next, cache)).toMatchObject({ indent: 4, gridDropZones: true, saveData: true })
  })
})
