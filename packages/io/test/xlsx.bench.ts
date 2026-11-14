// Reading and writing a large workbook: 100,000 rows × 10 columns (1M cells: numbers, text, formulas).
import { describe, expect, it } from "vitest"
import { Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"

function large(): Workbook {
  const wb = new Workbook(["Data"])
  const s = wb.sheets[0]
  for (let r = 0; r < 100_000; r++) {
    for (let c = 0; c < 8; c++) s.cells.set(r, c, { value: c % 2 ? r * 1.5 + c : `Text ${r % 1000}-${c}` })
    s.cells.set(r, 8, { value: r, input: `=B${r + 1}*2` })
    s.cells.set(r, 9, { value: r % 3 === 0 })
  }
  return wb
}

describe("xlsx, 1M cells", () => {
  it("saves and opens within budget", async () => {
    const wb = large()
    let t = performance.now()
    const bytes = await saveXlsx(wb)
    const save = performance.now() - t
    t = performance.now()
    const read = await openXlsx(bytes)
    const open = performance.now() - t
    console.log(`save ${save.toFixed(0)} ms, open ${open.toFixed(0)} ms, ${(bytes.length / 1e6).toFixed(1)} MB`)
    expect(read.workbook.sheets[0].cells.valueCount).toBe(1_000_000)
    expect(save).toBeLessThan(15_000)
    expect(open).toBeLessThan(15_000)
  })
})
