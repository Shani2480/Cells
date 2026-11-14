import { describe, expect, it } from "vitest"
import { MAX_ROWS, parseRange } from "../src/address"
import { encodeTSV } from "../src/clipboard"
import { readBlock } from "../src/edit"
import { formatValue } from "../src/format"
import { dataEdge } from "../src/navigation"
import type { CellWrite } from "../src/ops"
import { selectionStats } from "../src/stats"
import { Workbook } from "../src/workbook"

const ROWS = 100_000
const COLS = 20

function time<T>(label: string, run: () => T, budgetMs: number): T {
  const start = performance.now()
  const result = run()
  const ms = performance.now() - start
  console.log(`${label.padEnd(44)} ${ms.toFixed(1).padStart(8)} ms  (budget ${budgetMs} ms)`)
  expect(ms, label).toBeLessThan(budgetMs)
  return result
}

/** Phase 1 scale budgets (PLAN.md): 100,000 filled rows × 20 columns. */
describe("100,000 × 20 sheet", () => {
  const wb = new Workbook()
  const sheet = wb.sheets[0]

  it("loads, navigates, edits and undoes within budget", () => {
    time(
      "fill 2,000,000 cells (one transaction)",
      () => {
        const cells: CellWrite[] = []
        for (let row = 0; row < ROWS; row++) {
          for (let col = 0; col < COLS; col++) {
            cells.push({ row, col, data: { value: col % 3 === 0 ? `Item ${row}` : row * COLS + col } })
          }
        }
        wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load", history: false })
      },
      3000,
    )

    time(
      "Ctrl+Down from A1 (run end)",
      () => expect(dataEdge(sheet, { row: 0, col: 0 }, "down").row).toBe(ROWS - 1),
      50,
    )
    time(
      "Ctrl+Down from the last row (sheet edge)",
      () => expect(dataEdge(sheet, { row: ROWS - 1, col: 0 }, "down").row).toBe(MAX_ROWS - 1),
      50,
    )
    time(
      "Ctrl+Right along a row",
      () => expect(dataEdge(sheet, { row: 50_000, col: 0 }, "right").col).toBe(COLS - 1),
      5,
    )

    time(
      "draw-path lookups: 60 rows × 20 cols × 60 frames",
      () => {
        let n = 0
        for (let frame = 0; frame < 60; frame++) {
          const top = frame * 1500
          for (let row = top; row < top + 60; row++) {
            for (let col = 0; col < COLS; col++) {
              const data = sheet.cells.get(row, col)
              n += (formatValue(data?.value) ?? "").length
            }
          }
        }
        expect(n).toBeGreaterThan(0)
      },
      100,
    )

    time(
      "status bar over column B (whole column)",
      () => expect(selectionStats(sheet, [parseRange("B:B")!]).count).toBe(ROWS),
      300,
    )

    time(
      "insert a row at the top",
      () => wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 1 }], { label: "Insert" }),
      250,
    )
    time("undo it", () => wb.undo(), 250)
    expect(sheet.cells.value(0, 0)).toBe("Item 0")

    time(
      "copy 10,000 × 20 as TSV",
      () => {
        const block = readBlock(sheet, parseRange("A1:T10000")!)
        const text = encodeTSV(block.cells.map((row) => row.map((cell) => formatValue(cell?.value) ?? "")))
        expect(text.length).toBeGreaterThan(100_000)
      },
      500,
    )

    time(
      "type into 1,000 cells, 1,000 undo steps",
      () => {
        wb.undoLimit = 1000
        for (let i = 0; i < 1000; i++) {
          wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: i, col: 1, data: { value: -i } }] }], {
            label: "Typing",
          })
        }
        for (let i = 0; i < 1000; i++) wb.undo()
        expect(sheet.cells.value(5, 1)).toBe(5 * COLS + 1)
      },
      200,
    )
  })
})
