import { describe, expect, it } from "vitest"
import { pivotLayout } from "./engine"
import {
  fieldFilter,
  pivotChoices,
  pivotFilterTarget,
  withFieldCleared,
  withFieldFilter,
  withPageItems,
  withPagePick,
  withShownItems,
  withSort,
} from "./filter"
import { newPivotField, type PivotCache, type PivotTable } from "./model"

/** Sales by Region (rows), Product as a report filter, at A4 below it. */
function setup(): { cache: PivotCache; table: PivotTable } {
  const records = [
    ["North", "Apples", 10],
    ["South", "Pears", 20],
    ["East", "Apples", 5],
    ["North", "Pears", 7],
  ]
  const cache: PivotCache = {
    id: 1,
    source: { kind: "range", sheet: "Data", ref: "A1:C5" },
    fields: [
      { name: "Region", items: ["North", "South", "East"] },
      { name: "Product", items: ["Apples", "Pears"] },
      { name: "Sales", items: null },
    ],
    records,
    refreshOnLoad: false,
  }
  const fields = [newPivotField(3), newPivotField(2), newPivotField(0)]
  fields[0].axis = "row"
  fields[1].axis = "page"
  fields[2].dataField = true
  const table: PivotTable = {
    name: "PivotTable1",
    cacheId: 1,
    at: { row: 3, col: 0 },
    location: { r0: 3, c0: 0, r1: 7, c1: 1 },
    fields,
    rows: [0],
    cols: [],
    pages: [{ field: 1, item: null }],
    values: [{ name: "Sum of Sales", field: 2, fn: "sum" }],
    rowGrandTotals: true,
    colGrandTotals: true,
    compact: true,
    outline: true,
    compactData: true,
    captions: { data: "Values" },
    showError: false,
    showMissing: true,
    style: { name: null, rowHeaders: true, colHeaders: true, rowStripes: false, colStripes: false, lastColumn: true },
  }
  return { cache, table }
}

const cellAt = (t: PivotTable, c: PivotCache, row: number, col: number) =>
  pivotLayout(t, c).cells.find((x) => x.row === row && x.col === col)
const column = (t: PivotTable, c: PivotCache) =>
  pivotLayout(t, c)
    .cells.filter((x) => x.col <= 1 && x.row >= 3)
    .sort((a, b) => a.row - b.row || a.col - b.col)
    .map((x) => x.value)

describe("PivotTable filter dropdowns", () => {
  it("are on Row Labels and on the report filter's value", () => {
    const { cache, table } = setup()
    expect(pivotFilterTarget(table, cache, cellAt(table, cache, 3, 0))).toEqual({ kind: "row", fields: [0] })
    expect(pivotFilterTarget(table, cache, cellAt(table, cache, 1, 1))).toEqual({ kind: "page", fields: [1] })
    expect(pivotFilterTarget(table, cache, cellAt(table, cache, 1, 0))).toBeNull()
    expect(pivotFilterTarget(table, cache, cellAt(table, cache, 4, 0))).toBeNull()
  })

  it("outline form gives each row field's header its own button", () => {
    const { cache, table } = setup()
    const t: PivotTable = { ...table, rows: [0, 1], pages: [] }
    t.fields = t.fields.map((f, i) => (i < 2 ? { ...f, axis: "row", compact: false } : f))
    expect(pivotFilterTarget(t, cache, cellAt(t, cache, 3, 0))).toEqual({ kind: "row", fields: [0] })
    expect(pivotFilterTarget(t, cache, cellAt(t, cache, 3, 1))).toEqual({ kind: "row", fields: [1] })
    // Compact: one Row Labels for both.
    const compact: PivotTable = { ...table, rows: [0, 1], pages: [] }
    compact.fields = compact.fields.map((f, i) => (i === 1 ? { ...f, axis: "row" } : f))
    expect(pivotFilterTarget(compact, cache, cellAt(compact, cache, 3, 0))).toEqual({ kind: "row", fields: [0, 1] })
  })

  it("unticking items hides their rows and keeps the totals to what's shown", () => {
    const { cache, table } = setup()
    expect(pivotChoices(table, cache, 0).items.map((i) => [i.label, i.checked])).toEqual([
      ["North", true],
      ["South", true],
      ["East", true],
    ])
    const t = withShownItems(table, cache, 0, new Set([0, 2]))
    expect(t.fields[0].items.map((i) => Boolean(i.hidden))).toEqual([false, true, false])
    expect(column(t, cache).slice(-8)).toEqual([
      "Row Labels",
      "Sum of Sales",
      "North",
      17,
      "East",
      5,
      "Grand Total",
      22,
    ])
    expect(pivotChoices(t, cache, 0).filtered).toBe(true)
    // Nothing ticked changes nothing; Clear Filter shows all again.
    expect(withShownItems(t, cache, 0, new Set())).toBe(t)
    expect(withShownItems(t, cache, 0, null).fields[0].items.some((i) => i.hidden)).toBe(false)
    expect(withShownItems(table, cache, 0, null)).toBe(table)
  })

  it("sorts keep the field's items in the order shown", () => {
    const { cache, table } = setup()
    const t = withSort(table, cache, 0, "descending")
    expect(t.fields[0].sort).toBe("descending")
    expect(t.fields[0].items.map((i) => i.x)).toEqual([1, 0, 2])
    expect(pivotChoices(t, cache, 0).items.map((i) => i.label)).toEqual(["South", "North", "East"])
    expect(column(t, cache).slice(2, 8)).toEqual(["South", 20, "North", 17, "East", 5])
    expect(withSort(t, cache, 0, "descending")).toBe(t)
  })

  it("a report filter picks one item, or several with Select Multiple Items", () => {
    const { cache, table } = setup()
    expect(cellAt(table, cache, 1, 1)?.value).toBe("(All)")
    const one = withPagePick(table, cache, 1, 1)
    expect(one.pages[0].item).toBe(1)
    expect(cellAt(one, cache, 1, 1)?.value).toBe("Pears")
    expect(column(one, cache).slice(-6)).toEqual(["North", 7, "South", 20, "Grand Total", 27])
    expect(pivotChoices(one, cache, 1)).toMatchObject({ picked: 1, filtered: true })
    expect(withPagePick(one, cache, 1, null).pages[0].item).toBeNull()
    // Several: items hidden, the pick gone, multipleItemSelectionAllowed on.
    const some = withPageItems(one, cache, 1, new Set([0]))
    expect(some.pages[0].item).toBeNull()
    expect(some.fields[1].multipleItems).toBe(true)
    expect(some.fields[1].items.map((i) => Boolean(i.hidden))).toEqual([false, true])
    expect(cellAt(some, cache, 1, 1)?.value).toBe("Apples")
    const both = withPageItems(some, cache, 1, new Set([0, 1]))
    expect(cellAt(both, cache, 1, 1)?.value).toBe("(All)")
    // Back to one pick: the hidden items show again.
    const back = withPagePick(some, cache, 1, 0)
    expect(back.fields[1]).toMatchObject({ multipleItems: false })
    expect(back.fields[1].items.some((i) => i.hidden)).toBe(false)
  })

  it("a sorted report filter's pick follows its item", () => {
    const { cache, table } = setup()
    const one = withPagePick(table, cache, 1, 0)
    const sorted = withSort(one, cache, 1, "descending")
    expect(sorted.fields[1].items.map((i) => i.x)).toEqual([1, 0])
    expect(sorted.pages[0].item).toBe(1)
    expect(cellAt(sorted, cache, 1, 1)?.value).toBe("Apples")
  })

  it("a Label Filter keeps the items whose label matches, wildcards and all", () => {
    const { cache, table } = setup()
    const t = withFieldFilter(table, 0, { kind: "label", op: "endsWith", value: "th" })
    expect(column(t, cache).slice(2)).toEqual(["North", 17, "South", 20, "Grand Total", 37])
    expect(column(withFieldFilter(table, 0, { kind: "label", op: "equal", value: "?ast" }), cache).slice(2, 4)).toEqual(
      ["East", 5],
    )
    const between = withFieldFilter(table, 0, { kind: "label", op: "between", value: "e", value2: "n" })
    expect(column(between, cache).slice(2, 4)).toEqual(["East", 5])
    expect(pivotChoices(t, cache, 0)).toMatchObject({ filtered: true, filter: { kind: "label", op: "endsWith" } })
  })

  it("a Value Filter and Top 10 judge each item by its total", () => {
    const { cache, table } = setup()
    const over = withFieldFilter(table, 0, { kind: "value", op: "greaterThan", measure: 0, value: 10 })
    expect(column(over, cache).slice(2)).toEqual(["North", 17, "South", 20, "Grand Total", 37])
    const top = withFieldFilter(table, 0, { kind: "top", by: "count", top: true, measure: 0, value: 1 })
    expect(column(top, cache).slice(2)).toEqual(["South", 20, "Grand Total", 20])
    const bottomSum = withFieldFilter(table, 0, { kind: "top", by: "sum", top: false, measure: 0, value: 20 })
    // East (5) then North (17) reach 20.
    expect(column(bottomSum, cache).slice(2)).toEqual(["North", 17, "East", 5, "Grand Total", 22])
  })

  it("a new filter replaces the field's old one; Clear Filter From clears items and filter", () => {
    const { cache, table } = setup()
    const a = withFieldFilter(table, 0, { kind: "label", op: "contains", value: "o" })
    const b = withFieldFilter(a, 0, { kind: "value", op: "lessThan", measure: 0, value: 18 })
    expect(b.filters).toHaveLength(1)
    expect(fieldFilter(b, 0)).toMatchObject({ kind: "value", id: 2 })
    expect(withFieldFilter(b, 0, null).filters).toBeUndefined()
    expect(withFieldFilter(table, 0, null)).toBe(table)
    const hidden = withShownItems(b, cache, 0, new Set([0]))
    const cleared = withFieldCleared(hidden, cache, 0)
    expect(cleared.filters).toBeUndefined()
    expect(cleared.fields[0].items.some((i) => i.hidden)).toBe(false)
  })
})
