import { describe, expect, it } from "vitest"
import { goalSeek, Workbook } from "../src"

function book(cells: [number, number, string | number][]) {
  const wb = new Workbook()
  wb.transact(
    [
      {
        kind: "setCells",
        sheet: 0,
        cells: cells.map(([row, col, v]) => ({
          row,
          col,
          data: typeof v === "string" ? { input: v, value: 0 } : { value: v },
        })),
      },
    ],
    { label: "Load", history: false },
  )
  wb.calculate(true)
  return wb
}

describe("Goal Seek", () => {
  it("finds the loan amount for a monthly payment, leaving the workbook as it was", () => {
    // B1 rate, B2 months, B3 amount, B4 payment.
    const wb = book([
      [0, 1, 0.05 / 12],
      [1, 1, 60],
      [2, 1, 10000],
      [3, 1, "=-PMT(B1,B2,B3)"],
    ])
    const r = goalSeek(wb, { sheet: 0, row: 3, col: 1 }, 500, { sheet: 0, row: 2, col: 1 })
    expect(r.found).toBe(true)
    expect(r.value).toBeCloseTo(26495.35, 1)
    expect(Math.abs((r.result ?? 0) - 500)).toBeLessThanOrEqual(0.001)
    expect(wb.sheets[0].cells.value(2, 1)).toBe(10000)
    expect(wb.canUndo).toBe(false)
  })

  it("solves a curve and says so when there's no answer", () => {
    const wb = book([
      [0, 0, 1],
      [0, 1, "=A1^2-2"],
      [1, 0, 1],
      [1, 1, "=A2^2+1"],
    ])
    const root = goalSeek(wb, { sheet: 0, row: 0, col: 1 }, 0, { sheet: 0, row: 0, col: 0 })
    expect(root.found).toBe(true)
    expect(root.value).toBeCloseTo(Math.SQRT2, 3)
    const none = goalSeek(wb, { sheet: 0, row: 1, col: 1 }, 0, { sheet: 0, row: 1, col: 0 })
    expect(none.found).toBe(false)
  })
})
