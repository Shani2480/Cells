// Grouping made in Cells against the groups Excel 365 made from the same data (fixtures/excel/pivot-group.xlsx):
// each cache there is ungrouped, then grouped again with the settings Excel used.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  derivedDateGroups,
  groupSettings,
  withItemsGrouped,
  withRangeGroup,
  withUngrouped,
  type PivotCache,
} from "@cells/core"
import { openXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}`, import.meta.url)))
const caches = async () => (await openXlsx(fixture("pivot-group.xlsx"), "pivot-group.xlsx")).workbook.pivotCaches

const ok = <T>(v: T | string): T => {
  if (typeof v === "string") throw new Error(v)
  return v
}

describe("grouping as Excel groups", () => {
  it("dates by months, by years, quarters and months, by 7 days, and by quarters and years", async () => {
    const all = await caches()
    const cases: [number, string[], number][] = [
      [0, ["months"], 1],
      [1, ["months", "quarters", "years"], 1],
      [2, ["days"], 7],
      [7, ["quarters", "years"], 1],
    ]
    for (const [i, parts, by] of cases) {
      const excel = all[i]
      const plain = ok(withUngrouped(excel, 2))!.cache
      expect(derivedDateGroups(plain, 2)).toEqual([])
      expect(groupSettings(plain, 2)).toMatchObject({ start: null, end: null, parts: ["months"] })
      const made = ok(withRangeGroup(plain, 2, { start: null, end: null, by, parts: parts as never }))
      expect(made.fields, `cache ${i}`).toEqual(excel.fields)
      expect(groupSettings(made, 2)).toMatchObject({ parts, by })
    }
  })

  it("numbers in ranges, given or automatic", async () => {
    const all = await caches()
    const cases: [number, number, number | null, number | null, number][] = [
      [3, 3, 1, 17, 5],
      [4, 4, 0, 500, 100],
      [5, 3, null, null, 4],
    ]
    for (const [i, f, start, end, by] of cases) {
      const excel = all[i]
      const plain = withUngrouped(excel, f)!.cache
      expect(plain.fields[f].group).toBeUndefined()
      const made = ok(withRangeGroup(plain, f, { start, end, by }))
      expect(made.fields, `cache ${i}`).toEqual(excel.fields)
    }
  })

  it("Group Selection makes Product2 with Group1, then Product3; Ungroup takes them away again", async () => {
    const excel = (await caches())[6]
    const plain = withUngrouped(excel, 6)!
    expect(plain.removed).toBe(true)
    expect(plain.cache.fields).toHaveLength(6)
    const made = ok(withItemsGrouped(plain.cache, 1, [0, 2]))
    expect(made.field).toBe(6)
    expect(made.cache.fields).toEqual(excel.fields)
    // One item alone can't be grouped, as in Excel.
    expect(withItemsGrouped(made.cache, 1, [1])).toBe("Cannot group that selection.")
    // Product2's Plums and Group1 together: Excel makes Product3 on Product itself, with Group2.
    const three = ok(withItemsGrouped(made.cache, 6, [0, 1]))
    expect(three.field).toBe(7)
    expect(three.cache.fields[7]).toEqual({
      name: "Product3",
      items: null,
      group: { kind: "discrete", base: 1, map: [0, 0, 0], items: ["Group2"] },
    })
    // Ungroup Group2: Product3 goes, Product2 stays.
    const back = withUngrouped(three.cache, 7, [0])!
    expect(back).toMatchObject({ removed: true, base: 6 })
    expect(back.cache.fields).toEqual(excel.fields)
    // Splitting Group1 of Product2 leaves Apples, Plums and Pears: Product2 goes.
    expect(withUngrouped(made.cache, 6, [1])).toMatchObject({ removed: true, base: 1 })
  })

  it("refuses text and an end before the start", async () => {
    const c: PivotCache = (await caches())[6]
    expect(withRangeGroup(c, 0, { start: null, end: null, by: 10 })).toBe("Cannot group that selection.")
    const units = withUngrouped((await caches())[3], 3)!.cache
    expect(withRangeGroup(units, 3, { start: 10, end: 5, by: 1 })).toBe(
      "The ending value must be greater than the starting value.",
    )
  })
})
