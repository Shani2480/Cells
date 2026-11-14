import { describe, expect, it } from "vitest"
import {
  filterChoices,
  filterRowEntries,
  shiftAutoFilter,
  valuesFilter,
  snapshotSheet,
  Workbook,
  type AutoFilter,
} from "../src"

function book() {
  const wb = new Workbook()
  const s = wb.sheets[0]
  s.cells.set(0, 0, { value: "Fruit" })
  s.cells.set(0, 1, { value: "Qty" })
  ;["apple", "pear", "apple", "plum"].forEach((v, i) => {
    s.cells.set(i + 1, 0, { value: v })
    s.cells.set(i + 1, 1, { value: i * 10 })
  })
  return { wb, s }
}

describe("the autoFilter op", () => {
  it("sets the filter and hides its rows in one undo step", () => {
    const { wb, s } = book()
    const filter: AutoFilter = {
      range: { r0: 0, r1: 4, c0: 0, c1: 1 },
      columns: { 0: { kind: "values", values: ["apple"] } },
    }
    s.autoFilter = filter
    const rows = filterRowEntries(s, wb.styles, filter, 45000)
    s.autoFilter = null
    wb.transact([{ kind: "autoFilter", sheet: 0, filter, rows }], { label: "Filter" })
    expect(s.autoFilter?.columns[0]).toEqual({ kind: "values", values: ["apple"] })
    expect([1, 2, 3, 4].map((r) => s.rows.isHidden(r))).toEqual([false, true, false, true])
    wb.undo()
    expect(s.autoFilter).toBeNull()
    expect([1, 2, 3, 4].map((r) => s.rows.isHidden(r))).toEqual([false, false, false, false])
    wb.redo()
    expect(s.rows.isHidden(2)).toBe(true)
  })

  it("follows inserted and deleted rows and columns, and comes back on undo", () => {
    const { wb, s } = book()
    wb.transact(
      [
        {
          kind: "autoFilter",
          sheet: 0,
          filter: {
            range: { r0: 0, r1: 4, c0: 0, c1: 1 },
            columns: { 1: { kind: "top10", top: true, percent: false, count: 2 } },
          },
        },
      ],
      { label: "Filter" },
    )
    wb.transact([{ kind: "insertRows", sheet: 0, at: 2, count: 2 }], { label: "Insert" })
    expect(s.autoFilter?.range).toEqual({ r0: 0, r1: 6, c0: 0, c1: 1 })
    wb.transact([{ kind: "insertCols", sheet: 0, at: 0, count: 1 }], { label: "Insert" })
    expect(s.autoFilter?.range).toEqual({ r0: 0, r1: 6, c0: 1, c1: 2 })
    expect(s.autoFilter?.columns[1]?.kind).toBe("top10")
    wb.transact([{ kind: "deleteCols", sheet: 0, at: 2, count: 1 }], { label: "Delete" })
    expect(s.autoFilter?.range).toEqual({ r0: 0, r1: 6, c0: 1, c1: 1 })
    expect(s.autoFilter?.columns).toEqual({})
    wb.undo()
    expect(s.autoFilter?.columns[1]?.kind).toBe("top10")
    wb.undo()
    wb.undo()
    expect(s.autoFilter?.range).toEqual({ r0: 0, r1: 4, c0: 0, c1: 1 })
  })

  it("moves filters of columns after an inserted column inside the range", () => {
    const f: AutoFilter = {
      range: { r0: 0, r1: 9, c0: 2, c1: 5 },
      columns: { 0: { kind: "dynamic", type: "today" }, 2: { kind: "dynamic", type: "aboveAverage" } },
    }
    const next = shiftAutoFilter(f, "cols", 3, 2)
    expect(next?.range).toEqual({ r0: 0, r1: 9, c0: 2, c1: 7 })
    expect(Object.keys(next?.columns ?? {})).toEqual(["0", "4"])
    expect(shiftAutoFilter(f, "rows", 0, -10)).toBeNull()
  })

  it("is kept in a sheet snapshot", () => {
    const { s } = book()
    s.autoFilter = { range: { r0: 0, r1: 4, c0: 0, c1: 1 }, columns: {} }
    expect(snapshotSheet(s).autoFilter).toEqual(s.autoFilter)
  })
})

describe("a column's dropdown", () => {
  it("lists the values the other filters leave, ticked by this column's filter", () => {
    const { wb, s } = book()
    s.cells.set(5, 0, { value: "Apple" })
    s.cells.set(5, 1, { value: 99 })
    const filter: AutoFilter = {
      range: { r0: 0, r1: 6, c0: 0, c1: 1 },
      columns: {
        0: { kind: "values", values: ["apple"] },
        1: { kind: "custom", conditions: [{ op: "lt", value: "50" }] },
      },
    }
    const fruit = filterChoices(s, wb.styles, filter, 0, 45000)
    expect(fruit.kind).toBe("text")
    // "Apple" (99) is left out by Qty < 50; apple and Apple are one entry anyway.
    expect(fruit.values).toEqual([
      { text: "apple", checked: true },
      { text: "pear", checked: false },
      { text: "plum", checked: false },
    ])
    expect(fruit.blanks).toBeNull()
    const qty = filterChoices(s, wb.styles, filter, 1, 45000)
    expect(qty.kind).toBe("number")
    expect(qty.values.map((v) => `${v.text}:${v.checked}`)).toEqual(["0:true", "20:true", "99:false"])
    expect(qty.blanks).toBeNull()
  })

  it("keeps ticked dates at the widest whole level", () => {
    const d = (year: number, month: number, day: number, checked: boolean) => ({ year, month, day, checked })
    expect(
      valuesFilter({
        kind: "date",
        values: [{ text: "x", checked: false }],
        dates: [
          d(2024, 1, 5, true),
          d(2024, 2, 1, true),
          d(2025, 3, 1, true),
          d(2025, 3, 2, false),
          d(2025, 4, 9, true),
        ],
        blanks: { checked: true },
        cellColors: [],
        fontColors: [],
      }),
    ).toEqual({
      kind: "values",
      values: [],
      blanks: true,
      dates: [{ year: 2024 }, { year: 2025, month: 3, day: 1 }, { year: 2025, month: 4 }],
    })
  })
})
