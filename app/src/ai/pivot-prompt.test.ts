import { describe, expect, it } from "vitest"
import type { PivotCache } from "@cells/core"
import { pivotRequest, readPivotReply, recipeTitle } from "./pivot-prompt"

// 2025-01-10, 2025-04-12, 2026-01-15 as serials.
const cache: PivotCache = {
  id: 1,
  source: { kind: "range", sheet: "Data", ref: "A1:E5" },
  fields: [
    { name: "Order", items: null },
    { name: "Region", items: null },
    { name: "Date", items: null, dates: true, numFmt: "m/d/yyyy" },
    { name: "Units", items: null },
    { name: "Amount", items: null },
  ],
  records: [
    ["A-1001", "North", 45667, 4, 40],
    ["A-1002", "South", 45759, 2, 20],
    ["A-1003", "North", 46037, 5, 50],
    ["A-1004", "East", 45667, 1, 10],
  ],
  refreshOnLoad: false,
}

describe("Describe the PivotTable you want", () => {
  it("sends the columns, their kinds and a few values, never whole rows", () => {
    const r = pivotRequest(cache, "  sales by region, biggest first ", 4)
    expect(r.prompt).toBe("sales by region, biggest first")
    expect(r.context).toBe(
      [
        "The data has 4 rows and these columns:",
        "- Order: text, 4 different values (A-1001, A-1002, A-1003, A-1004)",
        "- Region: text, 3 different values (North, South, East)",
        "- Date: dates from 2025-01-10 to 2026-01-15",
        "- Units: numbers from 1 to 5",
        "- Amount: numbers from 10 to 50",
      ].join("\n"),
    )
  })

  it("reads the layout: columns by name in any case, summaries, filters, dates, sort by a value and Top N", () => {
    const reply = `Here it is:
{"title": "Amount by region and quarter", "rows": ["region"], "columns": ["Date"], "filters": ["Order"],
 "values": [{"field": "Units", "summary": "average"}, {"field": "Amount", "summary": "sum"}],
 "group_dates": {"field": "Date", "by": ["quarters", "years"]},
 "sort": {"field": "Region", "by": "Amount", "order": "descending"},
 "top": {"field": "Region", "count": 2, "by": "Sum of Amount"}}`
    expect(readPivotReply(reply, cache)).toEqual({
      recipe: {
        title: "Amount by region and quarter",
        rows: [1],
        cols: [2],
        pages: [0],
        values: [3, 4],
        fns: ["average", "sum"],
        dates: { field: 2, parts: ["quarters", "years"] },
        sort: { field: 1, by: 1, order: "descending" },
        top: { field: 1, count: 2, bottom: false, measure: 1 },
      },
    })
  })

  it("leaves out what doesn't fit: dates not on an axis, a sort on a column not shown, a label sort kept", () => {
    const reply = JSON.stringify({
      rows: ["Region", "Region"],
      columns: ["Region"],
      values: [{ field: "Amount" }],
      group_dates: { field: "Date", by: ["months"] },
      sort: { field: "Region", by: "label", order: "ascending" },
      top: { field: "Date", count: 3 },
    })
    expect(readPivotReply(reply, cache)).toEqual({
      recipe: {
        title: "Sum of Amount by Region",
        rows: [1],
        cols: [],
        values: [4],
        sort: { field: 1, by: "label", order: "ascending" },
      },
    })
  })

  it("says what went wrong: a column the data doesn't have, no columns, the model's own reason, no JSON", () => {
    expect(readPivotReply('{"rows": ["Country"], "values": [{"field": "Profit"}]}', cache)).toEqual({
      problem:
        "The AI used columns the data doesn't have: Country, Profit. Try again, or name the columns as they are written.",
    })
    expect(readPivotReply('{"rows": []}', cache)).toEqual({
      problem: "The AI didn't put any columns in the PivotTable. Try describing it another way.",
    })
    expect(readPivotReply('{"problem": "There is no profit column."}', cache)).toEqual({
      problem: "There is no profit column.",
    })
    expect(readPivotReply("Sorry, I can't.", cache)).toEqual({
      problem: "The AI didn't answer with a PivotTable. Try describing it another way.",
    })
  })

  it("titles a layout the model left untitled, as Recommended PivotTables does", () => {
    expect(recipeTitle(cache, { title: "", rows: [1], cols: [2], values: [4, 0] })).toBe(
      "Sum of Amount and Count of Order by Region and Date",
    )
  })
})
