import { describe, expect, it } from "vitest"
import { newSparklineGroup, parseRange, sparklinesFor, Workbook } from "../src"

function book() {
  const wb = new Workbook()
  wb.sheets[0].sparklines = [
    newSparklineGroup("line", [
      { row: 1, col: 5, source: "Sheet1!A2:E2" },
      { row: 2, col: 5, source: "Sheet1!A3:E3" },
    ]),
  ]
  return wb
}

describe("sparklines", () => {
  it("follow inserted and deleted rows and columns, with their data, and undo", () => {
    const wb = book()
    wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 1 }], { label: "x" })
    const g = () => wb.sheets[0].sparklines[0]
    expect(g().sparklines).toEqual([
      { row: 2, col: 5, source: "Sheet1!A3:E3" },
      { row: 3, col: 5, source: "Sheet1!A4:E4" },
    ])
    wb.transact([{ kind: "deleteCols", sheet: 0, at: 1, count: 1 }], { label: "x" })
    expect(g().sparklines[0]).toEqual({ row: 2, col: 4, source: "Sheet1!A3:D3" })
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 2, count: 1 }], { label: "x" })
    expect(g().sparklines).toEqual([{ row: 2, col: 4, source: "Sheet1!A3:D3" }])
    wb.undo()
    wb.undo()
    wb.undo()
    expect(g().sparklines).toEqual(book().sheets[0].sparklines[0].sparklines)
  })

  it("follow a renamed sheet", () => {
    const wb = book()
    wb.transact([{ kind: "renameSheet", sheet: 0, name: "Data 1" }], { label: "x" })
    expect(wb.sheets[0].sparklines[0].sparklines[0].source).toBe("'Data 1'!A2:E2")
  })

  it("Create Sparklines matches the data's lines to the location's cells", () => {
    const r = (t: string) => parseRange(t)!
    expect(sparklinesFor("Sheet1", r("A2:E4"), r("F2:F4"))).toEqual({
      sparklines: [
        { row: 1, col: 5, source: "Sheet1!A2:E2" },
        { row: 2, col: 5, source: "Sheet1!A3:E3" },
        { row: 3, col: 5, source: "Sheet1!A4:E4" },
      ],
    })
    expect(sparklinesFor("Sheet1", r("A1:C5"), r("A6:C6"))).toEqual({
      sparklines: [
        { row: 5, col: 0, source: "Sheet1!A1:A5" },
        { row: 5, col: 1, source: "Sheet1!B1:B5" },
        { row: 5, col: 2, source: "Sheet1!C1:C5" },
      ],
    })
    expect(sparklinesFor("Sheet1", r("A1:E1"), r("F1"))).toEqual({
      sparklines: [{ row: 0, col: 5, source: "Sheet1!A1:E1" }],
    })
    expect("error" in sparklinesFor("Sheet1", r("A1:E2"), r("F1"))).toBe(true)
    expect("error" in sparklinesFor("Sheet1", r("A1:E3"), r("F1:F2"))).toBe(true)
    expect("error" in sparklinesFor("Sheet1", r("A1:E3"), r("F1:G3"))).toBe(true)
  })
})
