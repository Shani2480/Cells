import { describe, expect, it } from "vitest"
import type { CellWrite } from "../src/ops"
import { Workbook } from "../src/workbook"

/**
 * Times `run` and checks it against the budget. With `repeat`, `run(i)` runs that many times and the best time
 * counts, so a busy machine (other processes, GC) doesn't fail a repeatable measurement.
 */
function time<T>(label: string, run: (i: number) => T, budgetMs: number, repeat = 1): T {
  let ms = Infinity
  let result!: T
  for (let i = 0; i < repeat; i++) {
    const start = performance.now()
    result = run(i)
    ms = Math.min(ms, performance.now() - start)
  }
  console.log(
    `${label.padEnd(52)} ${ms.toFixed(1).padStart(8)} ms  (budget ${budgetMs} ms${repeat > 1 ? `, best of ${repeat}` : ""})`,
  )
  expect(ms, label).toBeLessThan(budgetMs)
  return result
}

/** Phase 2 budgets (PLAN.md). */
describe("calculation budgets", () => {
  it("100,000-formula dependency chain recalculates in < 200 ms", () => {
    const wb = new Workbook()
    const cells: CellWrite[] = [{ row: 0, col: 0, data: { value: 1 } }]
    for (let r = 1; r < 100_000; r++) cells.push({ row: r, col: 0, data: { value: 0, input: `=A${r}+1` } })
    time(
      "enter 100,000 chained formulas",
      () => wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" }),
      5000,
    )
    time(
      "recalculate the chain after editing A1",
      (i) =>
        wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: { value: 2 + i } }] }], {
          label: "Edit",
        }),
      200,
      3,
    )
    expect(wb.sheets[0].cells.value(99_999, 0)).toBe(100_003)
  })

  it("inserting a row above 100,000 formulas rewrites them in < 2 s", () => {
    const wb = new Workbook()
    const cells: CellWrite[] = [{ row: 0, col: 0, data: { value: 1 } }]
    for (let r = 1; r < 100_000; r++) cells.push({ row: r, col: 0, data: { value: 0, input: `=A${r}+1` } })
    wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" })
    time(
      "insert a row above the chain (rewrite + rebuild)",
      () => wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 1 }], { label: "Insert" }),
      2000,
    )
    expect(wb.sheets[0].cells.get(100_000, 0)?.input).toBe("=A100000+1")
    time("undo the insert", () => wb.undo(), 3000)
    expect(wb.sheets[0].cells.get(99_999, 0)?.input).toBe("=A99999+1")
  })

  it("50,000 VLOOKUPs over 10,000 rows recalculate in < 1 s", () => {
    const wb = new Workbook()
    const table: CellWrite[] = []
    for (let r = 0; r < 10_000; r++) {
      table.push({ row: r, col: 0, data: { value: `K${r}` } })
      table.push({ row: r, col: 1, data: { value: r * 2 } })
    }
    wb.transact([{ kind: "setCells", sheet: 0, cells: table }], { label: "Table" })
    const lookups: CellWrite[] = []
    for (let r = 0; r < 50_000; r++) {
      lookups.push({ row: r, col: 3, data: { value: `K${(r * 7) % 10_000}` } })
      lookups.push({ row: r, col: 4, data: { value: 0, input: `=VLOOKUP(D${r + 1},$A$1:$B$10000,2,FALSE)` } })
    }
    time(
      "enter 50,000 VLOOKUPs",
      () => wb.transact([{ kind: "setCells", sheet: 0, cells: lookups }], { label: "Lookups" }),
      5000,
    )
    expect(wb.sheets[0].cells.value(1, 4)).toBe(14)
    time(
      "change the table: all 50,000 recalculate",
      () =>
        wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 5, col: 1, data: { value: -1 } }] }], {
          label: "Edit",
        }),
      1000,
    )
    // Row r looks up K((7r) mod 10000); the first row looking up K5 shows the new value.
    const r = Array.from({ length: 50_000 }, (_, i) => i).find((i) => (i * 7) % 10_000 === 5)!
    expect(wb.sheets[0].cells.value(r, 4)).toBe(-1)
  })

  it("SUM of a whole column under 50,000 values recalculates quickly", () => {
    const wb = new Workbook()
    const cells: CellWrite[] = []
    for (let r = 0; r < 50_000; r++) cells.push({ row: r, col: 0, data: { value: r } })
    cells.push({ row: 0, col: 1, data: { value: 0, input: "=SUM(A:A)" } })
    wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" })
    time(
      "edit one value under SUM(A:A)",
      () =>
        wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 3, col: 0, data: { value: 1000 } }] }], {
          label: "Edit",
        }),
      100,
    )
  })
})
