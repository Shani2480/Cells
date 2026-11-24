import { describe, expect, it } from "vitest"
import {
  CALC_ITEM_OTHER_FIELD,
  CALC_ITEM_NAME_TAKEN,
  calcItemDisplay,
  calcItemProblem,
  calcItemsBlocked,
  withCalcItem,
  withoutCalcItem,
} from "./calc-items"
import { pivotLayout } from "./engine"
import { withSort } from "./filter"
import { newPivotTable, withFieldIn } from "./make"
import type { PivotCache, PivotFn, PivotTable } from "./model"

/** Excel's own runs had this data. */
const cache: PivotCache = {
  id: 1,
  source: { kind: "range", sheet: "Data", ref: "A1:C7" },
  fields: [
    { name: "Product", items: ["Pens", "Ink", "Paper", "Clips"] },
    { name: "Region", items: ["East", "West"] },
    { name: "Amount", items: null },
  ],
  records: [
    ["Pens", "East", 10],
    ["Ink", "West", 40],
    ["Paper", "East", 25],
    ["Pens", "West", 5],
    ["Clips", "East", 1],
    ["Ink", "East", 2],
  ],
  refreshOnLoad: false,
}

/** The table's cells as rows of text (blank for none), from its first row. */
function grid(t: PivotTable, k: PivotCache): string[] {
  const layout = pivotLayout(t, k)
  const out: string[][] = []
  for (const c of layout.cells) {
    const r = c.row - layout.range.r0
    const col = c.col - layout.range.c0
    out[r] ??= []
    const v = c.value
    out[r][col] = v === undefined || v === null ? "" : typeof v === "object" ? v.error : String(v)
  }
  return out.map((row) => Array.from(row, (v) => v ?? "").join("|"))
}

const table = (k: PivotCache, rows: number[], cols: number[], fn: PivotFn = "sum") => {
  let t = newPivotTable("P", k, { row: 0, col: 0 })
  for (const f of rows) t = withFieldIn(t, k, f, "rows")
  for (const f of cols) t = withFieldIn(t, k, f, "cols")
  t = withFieldIn(t, k, 2, "values")
  return { ...t, values: t.values.map((v) => ({ ...v, fn })) }
}

describe("calculated items", () => {
  it("work out each cell, and totals add them in, as Excel 365 did", () => {
    let k = withCalcItem(cache, 0, "Writing", "=Pens+Ink*2") as PivotCache
    k = withCalcItem(k, 1, "Diff", "= East-West") as PivotCache
    expect(k.calcItems).toEqual([
      { field: 0, x: 4, formula: "Product[Pens]+Product[Ink]*2" },
      { field: 1, x: 2, formula: "Region[East]-Region[West]" },
    ])
    expect(k.fields[0].items).toEqual(["Pens", "Ink", "Paper", "Clips", "Writing"])
    const t = withSort(table(k, [0], [1]), k, 0, "ascending")
    expect(grid(t, k).slice(2)).toEqual([
      "Clips|1||1|2",
      "Ink|2|40|-38|4",
      "Paper|25||25|50",
      "Pens|10|5|5|20",
      "Writing|14|85|-71|28",
      "Grand Total|52|130|-78|104",
    ])
  })

  it("show under every outer item; totals put them together with the value field's function", () => {
    const k = withCalcItem(cache, 0, "Writing", "=Pens+Ink*2") as PivotCache
    const rows = (fn: PivotFn) => {
      const t = withSort(table(k, [1, 0], [], fn), k, 0, "ascending")
      return grid({ ...t, compact: false, fields: t.fields.map((f) => ({ ...f, compact: false, outline: false })) }, k)
    }
    expect(rows("sum").slice(1)).toEqual([
      "East|Clips|1",
      "|Ink|2",
      "|Paper|25",
      "|Pens|10",
      "|Writing|14",
      "East Total||52",
      "West|Ink|40",
      "|Pens|5",
      "|Writing|85",
      "West Total||130",
      "Grand Total||182",
    ])
    const totals = (fn: PivotFn) => rows(fn).filter((r) => r.includes("Total") || r.includes("Writing"))
    expect(totals("count")).toEqual(["|Writing|3", "East Total||7", "|Writing|3", "West Total||5", "Grand Total||12"])
    expect(totals("max")).toEqual(["|Writing|14", "East Total||25", "|Writing|85", "West Total||85", "Grand Total||85"])
    expect(totals("min")).toEqual(["|Writing|14", "East Total||1", "|Writing|85", "West Total||5", "Grand Total||1"])
  })

  it("where two meet, the one last in solve order works it out", () => {
    let k = withCalcItem(cache, 0, "Both", "=Pens*Ink") as PivotCache
    k = withCalcItem(k, 1, "Diff", "=East-West") as PivotCache
    const cell = (kk: PivotCache) =>
      grid(table(kk, [0], [1]), kk)
        .find((r) => r.startsWith("Both"))!
        .split("|")[3]
    // Diff last: Both in East less Both in West (20 - 200); Both last: Pens' Diff times Ink's (5 × -38).
    expect(cell(k)).toBe("-180")
    expect(cell({ ...k, calcItems: [...k.calcItems!].reverse() })).toBe("-190")
  })

  it("work their formula out in the innermost cells only; totals add those up (Excel 365's)", () => {
    const k = withCalcItem(cache, 0, "Both", "=Pens*Ink") as PivotCache
    // Both: 20 in East, 200 in West, and 220 in all (not 15 × 42).
    expect(grid(withSort(table(k, [0], [1]), k, 0, "ascending"), k).slice(2)).toEqual([
      "Both|20|200|220",
      "Clips|1||1",
      "Ink|2|40|42",
      "Paper|25||25",
      "Pens|10|5|15",
      "Grand Total|58|245|303",
    ])
    // An outer calculated item's own total too.
    const outer = withSort(table(k, [0, 1], []), k, 0, "ascending")
    const flat = {
      ...outer,
      compact: false,
      fields: outer.fields.map((f) => ({ ...f, compact: false, outline: false })),
    }
    expect(grid(flat, k).slice(1, 4)).toEqual(["Both|East|20", "|West|200", "Both Total||220"])
  })

  it("read other calculated items, functions, places and quoted names", () => {
    let k = withCalcItem(cache, 0, "Writing", "=Pens+Ink*2") as PivotCache
    k = withCalcItem(k, 0, "Half", "=Writing/2+'Pens'") as PivotCache
    k = withCalcItem(k, 0, "Top", "=MAX(Pens,Ink)+Product[1]") as PivotCache
    expect(calcItemDisplay(k, k.calcItems![1])).toBe("=Writing/2+Pens")
    expect(k.calcItems![2].formula).toBe("MAX(Product[Pens],Product[Ink])+Product[1]")
    const east = grid(table(k, [0], []), k)
    // Writing 99 (15 + 42×2); Half 49.5 + 15; Top 42 + Pens (the first item).
    expect(east.slice(1, 8)).toEqual(["Pens|15", "Ink|42", "Paper|25", "Clips|1", "Writing|99", "Half|64.5", "Top|57"])
  })

  it("say why a formula or a name won't do", () => {
    expect(calcItemProblem(cache, 0, "=Region[East]")).toBe(CALC_ITEM_OTHER_FIELD)
    expect(calcItemProblem(cache, 0, "=Pens+A1")).not.toBeNull()
    expect(calcItemProblem(cache, 0, "=Pens+Nothing")).not.toBeNull()
    expect(withCalcItem(cache, 0, "Ink", "=Pens")).toBe(CALC_ITEM_NAME_TAKEN)
    const k = withCalcItem(cache, 0, "Writing", "=Pens") as PivotCache
    // Modify keeps its place.
    expect((withCalcItem(k, 0, "writing", "=Ink") as PivotCache).calcItems).toEqual([
      { field: 0, x: 4, formula: "Product[Ink]" },
    ])
    expect(withoutCalcItem(k, 0, 4).fields[0].items).toEqual(["Pens", "Ink", "Paper", "Clips"])
    expect(withoutCalcItem(k, 0, 4).calcItems).toBeUndefined()
    const avg = table(k, [0], [], "average")
    expect(calcItemsBlocked([avg])).toMatch(/Averages/)
  })
})

describe("calculated items on refresh", () => {
  it("stay after the field's items and keep working", async () => {
    const { refreshedCache, refreshedTable } = await import("./refresh")
    const k = withCalcItem(cache, 0, "Writing", "=Pens+Ink*2") as PivotCache
    const t = table(k, [0], [])
    const data = {
      fields: k.fields.slice(0, 3).map((f) => ({ name: f.name })),
      records: [...cache.records, ["Tape", "East", 3]],
    }
    const next = refreshedCache(k, data)
    expect(next.fields[0].items).toEqual(["Pens", "Ink", "Paper", "Clips", "Tape", "Writing"])
    expect(next.calcItems).toEqual([{ field: 0, x: 5, formula: "Product[Pens]+Product[Ink]*2" }])
    expect(grid(refreshedTable(t, k, next), next).slice(1)).toContain("Writing|99")
  })
})
