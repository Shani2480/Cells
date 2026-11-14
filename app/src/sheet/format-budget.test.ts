import { MAX_ROWS, type CellWrite } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"

/** Phase 3 budget (PLAN.md): formatting a 100k-row column takes under 300 ms, Undo included in the check. */
function best(run: (i: number) => void, repeat = 3): number {
  let ms = Infinity
  for (let i = 0; i < repeat; i++) {
    const start = performance.now()
    run(i)
    ms = Math.min(ms, performance.now() - start)
  }
  return ms
}

describe("formatting budget", () => {
  const c = new SheetController()
  const cells: CellWrite[] = []
  for (let r = 0; r < 100_000; r++) cells.push({ row: r, col: 0, data: { value: r * 1.5 } })
  c.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" })
  // Row heights fit by measuring text; a fixed measure keeps the test about formatting, not about fonts.
  c.rowFit = () => 20

  const select = (r1: number) =>
    c.setSelection(
      {
        ranges: [{ r0: 0, c0: 0, r1, c1: 0 }],
        active: { row: 0, col: 0 },
        activeRange: 0,
        anchor: { row: 0, col: 0 },
        focus: { row: r1, col: 0 },
      },
      null,
    )

  const cases: [string, () => void][] = [
    ["Bold", () => c.format.toggle("bold")],
    ["Currency", () => c.format.numberFormat('"$"#,##0.00')],
    ["Fill", () => c.format.fill("#FFFF00")],
    ["Font size 16 (row heights refit)", () => c.format.fontSize(16)],
    ["Borders", () => c.format.border("all")],
  ]

  for (const [name, run] of cases) {
    it(`${name} on A1:A100000 in < 300 ms`, () => {
      select(99_999)
      const ms = best(() => {
        run()
        c.wb.undo()
      })
      console.log(`${name.padEnd(36)} ${ms.toFixed(1).padStart(7)} ms`)
      expect(ms).toBeLessThan(300)
    })
  }

  it("Bold on the whole column A in < 300 ms", () => {
    select(MAX_ROWS - 1)
    const ms = best(() => {
      c.format.toggle("bold")
      c.wb.undo()
    })
    expect(ms).toBeLessThan(300)
  })
})
