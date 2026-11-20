import { describe, expect, it } from "vitest"
import { pivotLayout } from "./engine"
import { newPivotTable, withFieldIn } from "./make"
import type { PivotCache } from "./model"
import {
  PIVOT_FIELD_NAME_TAKEN,
  PIVOT_NAME_INVALID,
  pivotOptions,
  withFieldSettings,
  withPivotOptions,
  withValueSettings,
} from "./settings"
import { Sheet } from "../sheet"

const cache: PivotCache = {
  id: 1,
  source: { kind: "range", sheet: "Data", ref: "A1:C4" },
  fields: [
    { name: "Region", items: ["North", "South"] },
    { name: "Channel", items: ["Shop", "Online"] },
    { name: "Units", items: null },
  ],
  records: [
    ["North", "Shop", 4],
    ["South", "Online", 2],
    ["North", "Online", 6],
  ],
  refreshOnLoad: false,
}

const make = () => {
  let t = newPivotTable("PivotTable1", cache, { row: 0, col: 0 })
  t = withFieldIn(t, cache, 0, "rows")
  t = withFieldIn(t, cache, 2, "values")
  return t
}

const values = (t: ReturnType<typeof make>) =>
  pivotLayout(t, cache)
    .cells.filter((c) => c.col === 1)
    .map((c) => c.value)

describe("Value Field Settings", () => {
  it("changing the function renames an automatic name, and keeps a typed one", () => {
    const t = withValueSettings(make(), cache, 0, { name: "Sum of Units", fn: "count" })
    if (typeof t === "string") throw new Error(t)
    expect(t.values[0]).toEqual({ name: "Count of Units", field: 2, fn: "count" })
    expect(values(t)).toEqual(["Count of Units", 2, 1, 3])
    const typed = withValueSettings(make(), cache, 0, { name: "Units sold", fn: "max", numFmt: "0.0" })
    expect(typed).toMatchObject({ values: [{ name: "Units sold", fn: "max", numFmt: "0.0" }] })
    // Empty: the automatic name again.
    expect(withValueSettings(make(), cache, 0, { name: "", fn: "average" })).toMatchObject({
      values: [{ name: "Average of Units" }],
    })
  })

  it("a name can't be a source field's or another value field's (a trailing space can)", () => {
    expect(withValueSettings(make(), cache, 0, { name: "units", fn: "sum" })).toBe(PIVOT_FIELD_NAME_TAKEN)
    expect(withValueSettings(make(), cache, 0, { name: "Units ", fn: "sum" })).toMatchObject({
      values: [{ name: "Units " }],
    })
    const two = withFieldIn(make(), cache, 2, "values")
    expect(withValueSettings(two, cache, 1, { name: "Sum of Units", fn: "sum" })).toBe(PIVOT_FIELD_NAME_TAKEN)
  })
})

describe("Field Settings", () => {
  const plain = {
    name: "",
    subtotals: ["default" as const],
    outline: true,
    compact: true,
    subtotalTop: true,
    repeatLabels: false,
    insertBlankRow: false,
    showAll: false,
  }
  it("names the field, unless another field or value has that name", () => {
    const t = withFieldSettings(make(), cache, 0, { ...plain, name: "Area" })
    expect((t as ReturnType<typeof make>).fields[0].name).toBe("Area")
    expect(withFieldSettings(make(), cache, 0, { ...plain, name: "channel" })).toBe(PIVOT_FIELD_NAME_TAKEN)
    expect(withFieldSettings(make(), cache, 0, { ...plain, name: "Sum of Units" })).toBe(PIVOT_FIELD_NAME_TAKEN)
    // Its own name back is no custom name.
    expect(
      (withFieldSettings(make(), cache, 0, { ...plain, name: "Region" }) as { fields: { name?: string }[] }).fields[0]
        .name,
    ).toBeUndefined()
  })

  it("subtotals, tabular form and a blank line after each item", () => {
    let t = withFieldIn(make(), cache, 1, "rows")
    t = withFieldSettings(t, cache, 0, {
      ...plain,
      outline: false,
      compact: false,
      insertBlankRow: true,
      subtotals: ["max"],
    }) as typeof t
    expect(t.fields[0]).toMatchObject({ outline: false, compact: false, insertBlankRow: true, subtotals: ["max"] })
  })
})

describe("PivotTable Options", () => {
  it("renames, sets captions and flags; the name is needed and unique on its sheet", () => {
    const sheet = new Sheet("S")
    const t = make()
    const other = { ...make(), name: "Other" }
    sheet.pivots = [t, other]
    const now = pivotOptions(t, cache)
    expect(now).toMatchObject({ name: "PivotTable1", autoFit: true, valuesRow: false, enableDrill: true })
    expect(withPivotOptions(t, sheet, { ...now, name: " " })).toBe(PIVOT_NAME_INVALID)
    expect(withPivotOptions(t, sheet, { ...now, name: "other" })).toMatch(/already exists/)
    const next = withPivotOptions(t, sheet, {
      ...now,
      name: "Sales",
      missingCaption: "-",
      autoFit: false,
      enableDrill: false,
      colGrandTotals: false,
      valuesRow: true,
    })
    expect(next).toMatchObject({ name: "Sales", captions: { missing: "-" }, autoFit: false, enableDrill: false })
    expect((next as typeof t).hideValuesRow).toBeUndefined()
    const back = withPivotOptions(next as typeof t, sheet, { ...now, name: "Sales" }) as typeof t
    expect(back.autoFit).toBeUndefined()
    expect(back.captions.missing).toBeUndefined()
  })
})
