import { describe, expect, it } from "vitest"
import { pivotLayout } from "./engine"
import type { PivotCache } from "./model"
import { pivotFromRecipe, recommendedPivots } from "./recommend"

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
})
