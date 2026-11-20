import { describe, expect, it } from "vitest"
import { calcFieldInputs, calcFormulaProblem, withCalcField, withoutCalcField, CALC_FORMULA_REFS } from "./calculated"
import { pivotLayout } from "./engine"
import { newPivotTable, withFieldIn } from "./make"
import { refreshedCache, refreshedTable } from "./refresh"
import type { PivotCache } from "./model"

const cache: PivotCache = {
  id: 1,
  source: { kind: "range", sheet: "Data", ref: "A1:D5" },
  fields: [
    { name: "Region", items: ["North", "South"] },
    { name: "Units", items: null },
    { name: "Amount", items: null },
    { name: "Unit Cost", items: null },
  ],
  records: [
    ["North", 2, 20, 1],
    ["North", 3, 40, 1],
    ["South", 0, 15, 0],
    ["South", 5, null, 2],
  ],
  refreshOnLoad: false,
}

/** The values column of a PivotTable by Region of value field `f`. */
function column(k: PivotCache, f: number, fn: "sum" | "count" | "average" = "sum") {
  let t = withFieldIn(newPivotTable("P", k, { row: 0, col: 0 }), k, 0, "rows")
  t = withFieldIn(t, k, f, "values")
  t = { ...t, values: t.values.map((v) => ({ ...v, fn })) }
  return pivotLayout(t, k)
    .cells.filter((c) => c.col === 1 && c.row > 0)
    .map((c) => c.value)
}

describe("calculated fields", () => {
  it("work their formula out on each cell's sums, as Excel does", () => {
    const k = withCalcField(cache, "Price", "=Amount/Units") as PivotCache
    expect(k.fields[4]).toEqual({ name: "Price", items: null, formula: "Amount/Units" })
    // North 60/5, South 15/5, total 75/10.
    expect(column(k, 4)).toEqual([12, 3, 7.5])
    // Count and Average too work the formula out on the sums (Excel's Count of Price is the price).
    expect(column(k, 4, "count")).toEqual([12, 3, 7.5])
    expect(column(k, 4, "average")).toEqual([12, 3, 7.5])
  })

  it("reads quoted names, other calculated fields and functions; a zero divisor is #DIV/0!", () => {
    let k = withCalcField(cache, "Per Cost", "=Amount/'Unit Cost'") as PivotCache
    k = withCalcField(k, "Big", "=IF(Units>4,Amount*2,0)") as PivotCache
    k = withCalcField(k, "Twice", "= Big + 'Per Cost'") as PivotCache
    expect(column(k, 4)).toEqual([30, 7.5, 75 / 4])
    expect(column(k, 5)).toEqual([120, 30, 150])
    expect(column(k, 6)).toEqual([150, 37.5, 150 + 75 / 4])
    expect(calcFieldInputs(k, 6).sort()).toEqual([1, 2, 3])
    const zero = withCalcField(cache, "Per Unit", "=Amount/Units") as PivotCache
    const t = withFieldIn(withFieldIn(newPivotTable("P", zero, { row: 0, col: 0 }), zero, 1, "rows"), zero, 4, "values")
    const cells = pivotLayout(t, zero).cells
    const row = cells.find((c) => c.col === 0 && c.value === 0)!.row
    expect(cells.find((c) => c.row === row && c.col === 1)?.value).toEqual({ error: "#DIV/0!" })
  })

  it("refuses what Excel refuses, and Modify keeps the field in place", () => {
    expect(calcFormulaProblem(cache, "=Amount+A1")).toBe(CALC_FORMULA_REFS)
    expect(calcFormulaProblem(cache, "=Amount+Nope")).toBe(CALC_FORMULA_REFS)
    expect(calcFormulaProblem(cache, "=SUM({1,2})")).toBe(CALC_FORMULA_REFS)
    expect(calcFormulaProblem(cache, "=Amount+")).toBe("There's a problem with this formula.")
    expect(calcFormulaProblem(cache, '=IF(Units>1,"big",0)')).toBeNull()
    expect(withCalcField(cache, "units", "=1")).toBe("PivotTable field name already exists.")
    const k = withCalcField(cache, "Price", "=Amount/Units") as PivotCache
    const m = withCalcField(k, "price", "=Amount*2") as PivotCache
    expect(m.fields.map((f) => f.formula)).toEqual([undefined, undefined, undefined, undefined, "Amount*2"])
    expect(withoutCalcField(m, "PRICE").fields).toHaveLength(4)
  })

  it("stay through a refresh, after the source's columns", () => {
    const k = withCalcField(cache, "Price", "=Amount/Units") as PivotCache
    let t = withFieldIn(withFieldIn(newPivotTable("P", k, { row: 0, col: 0 }), k, 0, "rows"), k, 4, "values")
    const data = {
      fields: [...cache.fields.slice(0, 3), { name: "Channel", items: null }, cache.fields[3]],
      records: cache.records.map((r) => [r[0], r[1], r[2], "Shop", r[3]]),
    }
    const next = refreshedCache(k, data)
    expect(next.fields.map((f) => f.name)).toEqual(["Region", "Units", "Amount", "Channel", "Unit Cost", "Price"])
    t = refreshedTable(t, k, next)
    expect(t.values[0].field).toBe(5)
    expect(
      pivotLayout(t, next)
        .cells.filter((c) => c.col === 1 && c.row > 0)
        .map((c) => c.value),
    ).toEqual([12, 3, 7.5])
  })

  it("can only be values", () => {
    const k = withCalcField(cache, "Price", "=Amount/Units") as PivotCache
    const t = newPivotTable("P", k, { row: 0, col: 0 })
    expect(withFieldIn(t, k, 4, "rows")).toBe(t)
    expect(withFieldIn(t, k, 4, "filters")).toBe(t)
  })
})
