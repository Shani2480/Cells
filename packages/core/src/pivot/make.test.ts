import { describe, expect, it } from "vitest"
import { pivotLayout } from "./engine"
import { withFieldFilter } from "./filter"
import {
  newPivotTable,
  nextPivotName,
  pivotFieldAreas,
  pivotIsEmpty,
  withFieldIn,
  withFieldRemoved,
  withFieldTicked,
  withoutField,
  withValueMoved,
} from "./make"
import { VALUES_FIELD, type PivotCache } from "./model"
import { Sheet } from "../sheet"

const cache: PivotCache = {
  id: 3,
  source: { kind: "range", sheet: "Data", ref: "A1:D4" },
  fields: [
    { name: "Region", items: ["North", "South"] },
    { name: "Day", items: [45658, 45659], dates: true, numFmt: "m/d/yyyy" },
    { name: "Units", items: null },
    { name: "Note", items: ["a", null] },
  ],
  records: [
    ["North", 45658, 4, "a"],
    ["South", 45659, null, null],
    ["North", 45659, 6, "a"],
  ],
  refreshOnLoad: false,
}

describe("making a PivotTable", () => {
  it("starts empty, as Excel 365 makes one, with its placeholder's 18 by 3 cells", () => {
    const t = newPivotTable("PivotTable1", cache, { row: 2, col: 0 })
    expect(pivotIsEmpty(t)).toBe(true)
    expect(t.style.name).toBe("PivotStyleLight16")
    expect(t.hideValuesRow).toBe(true)
    const layout = pivotLayout(t, cache)
    expect(layout.range).toEqual({ r0: 2, c0: 0, r1: 19, c1: 2 })
    expect(layout.cells).toEqual([])
    expect(layout.rowItems).toEqual([])
  })

  it("names PivotTables PivotTable1, 2, ... across the workbook", () => {
    const a = new Sheet("A")
    const b = new Sheet("B")
    a.pivots = [newPivotTable("PivotTable1", cache, { row: 0, col: 0 })]
    b.pivots = [newPivotTable("pivottable2", cache, { row: 0, col: 0 })]
    expect(nextPivotName([a, b])).toBe("PivotTable3")
  })

  it("ticking puts numbers in Values (Sum) and text and dates in Rows", () => {
    let t = newPivotTable("PivotTable1", cache, { row: 2, col: 0 })
    t = withFieldTicked(t, cache, 0)
    t = withFieldTicked(t, cache, 2)
    t = withFieldTicked(t, cache, 1)
    expect(t.rows).toEqual([0, 1])
    expect(t.values).toEqual([{ name: "Sum of Units", field: 2, fn: "sum" }])
    expect(t.fields[0]).toMatchObject({ axis: "row", dataField: false })
    expect(t.fields[2]).toMatchObject({ axis: null, dataField: true })
    // Text in Values counts.
    t = withFieldIn(t, cache, 3, "values")
    expect(t.values[1]).toEqual({ name: "Count of Note", field: 3, fn: "count" })
    // Two value fields bring Σ Values into Columns; back to one takes it away.
    expect(t.cols).toEqual([VALUES_FIELD])
    t = withFieldIn(t, cache, 2, "values")
    expect(t.values[2].name).toBe("Sum of Units2")
    t = withoutField(t, 3, "values")
    t = withoutField(t, 2, "values", 1)
    expect(t.values.map((v) => v.name)).toEqual(["Sum of Units"])
    expect(t.cols).toEqual([])
  })

  it("a field is in one of Filters, Columns and Rows; moving it takes it out of the other", () => {
    let t = newPivotTable("PivotTable1", cache, { row: 2, col: 0 })
    t = withFieldIn(t, cache, 0, "rows")
    t = withFieldIn(t, cache, 1, "rows", 0)
    expect(t.rows).toEqual([1, 0])
    t = withFieldIn(t, cache, 0, "filters")
    expect(t.rows).toEqual([1])
    expect(t.pages).toEqual([{ field: 0, item: null }])
    expect(pivotFieldAreas(t, 0)).toEqual(["filters"])
    t = withFieldIn(t, cache, 0, "cols")
    expect(t.pages).toEqual([])
    expect(t.fields[0].axis).toBe("col")
    // Leaving the axes for good drops its label and value filters.
    t = withFieldFilter(t, 0, { kind: "label", op: "contains", value: "o" })
    t = withFieldRemoved(t, 0)
    expect(t.filters).toBeUndefined()
    expect(t.fields[0].axis).toBeNull()
  })

  it("value fields move among themselves", () => {
    let t = newPivotTable("PivotTable1", cache, { row: 2, col: 0 })
    t = withFieldIn(t, cache, 2, "values")
    t = withFieldIn(t, cache, 3, "values")
    t = withValueMoved(t, 1, 0)
    expect(t.values.map((v) => v.name)).toEqual(["Count of Note", "Sum of Units"])
  })
})
