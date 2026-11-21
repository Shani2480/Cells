import { describe, expect, it } from "vitest"
import { parseInput, Workbook } from "../src"

/** A sheet from rows of typed text (as entered in cells), from A1. */
function book(rows: string[][]) {
  const wb = new Workbook(["Data"])
  const cells = rows.flatMap((r, row) =>
    r.flatMap((text, col) => {
      const p = text === "" ? null : parseInput(text)
      return p ? [{ row, col, data: { value: p.value, input: p.input } }] : []
    }),
  )
  wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load", history: false })
  wb.calculate(true)
  return wb
}
const set = (wb: Workbook, row: number, col: number, text: string) => {
  const p = parseInput(text)!
  wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row, col, data: { value: p.value, input: p.input } }] }], {
    label: "Edit",
  })
  wb.calculate()
}
const value = (wb: Workbook, row: number, col: number) => wb.sheets[0].cells.value(row, col)

// Database A1:C5, criteria E1:E2 and F1:F2 (a computed criterion), results in H1:H3.
const ROWS = [
  ["Region", "Rep", "Sales", "", "Region", "Above average", "", '=DSUM(A1:C5,"Sales",E1:E2)'],
  ["East", "Kim", "100", "", "East", "=C2>AVERAGE($C$2:$C$5)", "", "=DSUM(A1:C5,3,F1:F2)"],
  ["West", "Lee", "250", "", "", "", "", '=DGET(A1:C5,"Rep",F1:F2)'],
  ["East", "Ana", "40"],
  ["North", "Max", "300"],
]

describe("database functions", () => {
  it("sum the records that meet plain and computed criteria", () => {
    const wb = book(ROWS)
    expect(value(wb, 0, 7)).toBe(140)
    expect(value(wb, 1, 7)).toBe(550)
    expect(value(wb, 2, 7)).toEqual({ error: "#NUM!" })
  })

  it("recalculate when the database or the criteria change", () => {
    const wb = book(ROWS)
    set(wb, 3, 2, "60")
    expect(value(wb, 0, 7)).toBe(160)
    set(wb, 1, 4, "North")
    expect(value(wb, 0, 7)).toBe(300)
    set(wb, 4, 2, "100")
    expect(value(wb, 1, 7)).toBe(250)
    expect(value(wb, 2, 7)).toBe("Lee")
  })
})
