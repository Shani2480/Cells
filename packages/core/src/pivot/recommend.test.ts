import { describe, expect, it } from "vitest"
import { pivotLayout } from "./engine"
import type { PivotCache, PivotTable } from "./model"
import { pivotFromRecipe, recommendedPivots, type PivotRecipe } from "./recommend"
import { withSort } from "./filter"
import { withoutField, withValueMoved } from "./make"

// 2025-01-10, 2025-02-12, 2026-01-15, 2026-03-02 as serials.
const days = [45667, 45700, 46037, 46083]
const cache: PivotCache = {
  id: 1,
  source: { kind: "range", sheet: "Data", ref: "A1:F9" },
  fields: [
    { name: "Order", items: null },
    { name: "Region", items: null },
    { name: "Channel", items: null },
    { name: "Date", items: null, dates: true, numFmt: "m/d/yyyy" },
    { name: "Units", items: null },
    { name: "Amount", items: null },
  ],
  records: [
    ["A1", "North", "Shop", days[0], 4, 40],
    ["A2", "South", "Online", days[1], 2, 20],
    ["A3", "North", "Online", days[2], 5, 50],
    ["A4", "East", "Shop", days[3], 1, 10],
    ["A5", "East", "Online", days[0], 3, 30],
    ["A6", "South", "Shop", days[1], 6, 60],
  ],
  refreshOnLoad: false,
}

describe("Recommended PivotTables", () => {
  it("sums the first number by each category, fewest items first, then by month, then the rest", () => {
    const titles = recommendedPivots(cache).map((r) => r.title)
    expect(titles).toEqual([
      "Sum of Units by Channel",
      "Sum of Units by Region",
      "Sum of Units by Date",
      "Sum of Amount by Date",
      "Sum of Amount by Channel",
      "Sum of Amount by Region",
      "Sum of Units and Sum of Amount by Channel",
      "Sum of Units by Region and Channel",
    ])
    // Order is different in every row: never a category.
    expect(titles.some((t) => t.includes("Order"))).toBe(false)
  })

  it("makes the PivotTable: dates grouped by years and months, the larger outermost", () => {
    const recipe = recommendedPivots(cache).find((r) => r.title === "Sum of Units by Date")!
    expect(recipe.dates).toEqual({ field: 3, parts: ["months", "years"] })
    const { table, cache: grouped } = pivotFromRecipe("P", cache, { row: 2, col: 0 }, recipe)
    expect(table.rows.map((f) => grouped.fields[f].name)).toEqual(["Years (Date)", "Months (Date)"])
    const labels = pivotLayout(table, grouped)
      .cells.filter((c) => c.col === 0 && typeof c.value === "string")
      .map((c) => c.value)
    expect(labels).toEqual(["Row Labels", "2025", "Jan", "Feb", "2026", "Jan", "Mar", "Grand Total"])
  })

  it("counts items when the data has no numbers", () => {
    const text: PivotCache = {
      ...cache,
      fields: cache.fields.slice(1, 3),
      records: cache.records.map((r) => r.slice(1, 3)),
    }
    expect(recommendedPivots(text).map((r) => r.title)).toEqual([
      "Count of Channel by Channel",
      "Count of Region by Region and Channel",
      "Count of Region by Region",
      "Count of Channel by Channel and Region",
    ])
  })

  /** The labels and numbers down the PivotTable's first two columns. */
  const column = (table: PivotTable, c: PivotCache) => {
    const layout = pivotLayout(table, c)
    const at = new Map(layout.cells.map((x) => [`${x.row},${x.col}`, x.value]))
    const out: [unknown, unknown][] = []
    for (let r = layout.range.r0; r <= layout.range.r1; r++)
      out.push([at.get(`${r},${layout.range.c0}`), at.get(`${r},${layout.range.c0 + 1}`)])
    return out
  }

  it("sorts rows by a value, biggest first, and keeps the sort through moved and removed values", () => {
    const recipe: PivotRecipe = {
      title: "",
      rows: [1],
      cols: [],
      values: [4, 5],
      sort: { field: 1, by: 1, order: "descending" },
    }
    const { table, cache: c } = pivotFromRecipe("P", cache, { row: 2, col: 0 }, recipe)
    expect(table.fields[1]).toMatchObject({ sort: "descending", sortBy: 1 })
    // Amount: North 90, South 80, East 40.
    expect(
      column(table, c)
        .slice(1, 4)
        .map((x) => x[0]),
    ).toEqual(["North", "South", "East"])
    // Units first is now value 1: the sort follows Amount.
    const moved = withValueMoved(table, 1, 0)
    expect(moved.fields[1].sortBy).toBe(0)
    // Amount gone: back to the labels' order.
    const gone = withoutField(moved, 5, "values")
    expect(gone.fields[1].sortBy).toBeUndefined()
    // Sort A to Z ends the sort by value.
    expect(withSort(table, c, 1, "ascending").fields[1].sortBy).toBeUndefined()
  })

  it("sorts by value within the item above, and smallest first", () => {
    const { table, cache: c } = pivotFromRecipe(
      "P",
      cache,
      { row: 2, col: 0 },
      {
        title: "",
        rows: [2, 1],
        cols: [],
        values: [5],
        sort: { field: 1, by: 0, order: "ascending" },
      },
    )
    // Channels in the data's order. Shop: East 10, North 40, South 60; Online: South 20, East 30, North 50.
    expect(column(table, c).map((x) => x[0])).toEqual([
      "Row Labels",
      "Shop",
      "East",
      "North",
      "South",
      "Online",
      "South",
      "East",
      "North",
      "Grand Total",
    ])
  })

  it("makes report filters, summaries, quarters on the columns and a Top N", () => {
    const { table, cache: c } = pivotFromRecipe(
      "P",
      cache,
      { row: 2, col: 0 },
      {
        title: "",
        rows: [1],
        cols: [3],
        values: [5],
        fns: ["average"],
        pages: [2],
        dates: { field: 3, parts: ["quarters", "years"] },
        top: { field: 1, count: 2, measure: 0 },
      },
    )
    expect(table.pages.map((p) => c.fields[p.field].name)).toEqual(["Channel"])
    expect(table.cols.map((f) => c.fields[f].name)).toEqual(["Years (Date)", "Quarters (Date)"])
    expect(table.values[0]).toMatchObject({ fn: "average", name: "Average of Amount" })
    // Average Amount: North 45, South 40, East 20: the top two stay.
    const labels = column(table, c).map((x) => x[0])
    expect(labels).toContain("North")
    expect(labels).toContain("South")
    expect(labels).not.toContain("East")
  })
})
