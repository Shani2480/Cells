// Writes the round-trip samples to tmp/xlsx when CELLS_WRITE_SAMPLES is set, for the Open XML validator and Excel.
import { readFileSync, writeFileSync, mkdirSync } from "node:fs"
import { it } from "vitest"
import { isError, type Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { samples } from "./samples"

it.runIf(process.env.CELLS_WRITE_SAMPLES)("writes samples", async () => {
  mkdirSync("tmp/xlsx", { recursive: true })
  for (const [name, make] of Object.entries(samples)) {
    const wb = make()
    writeFileSync(`tmp/xlsx/${name}.xlsx`, await saveXlsx(wb, { now: new Date(Date.UTC(2026, 0, 1)) }))
    writeFileSync(`tmp/xlsx/${name}.json`, JSON.stringify(expected(wb)))
  }
  // Excel's own files, saved again by Cells with their parts kept.
  for (const name of ["budget", "formats", "conditional", "layout", "protected", "features", "validation"]) {
    const bytes = new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}.xlsx`, import.meta.url)))
    const read = await openXlsx(bytes, `${name}.xlsx`)
    const out = await saveXlsx(read.workbook, { source: read.source, views: read.views, activeSheet: read.activeSheet })
    writeFileSync(`tmp/xlsx/${name}.resaved.xlsx`, out)
    writeFileSync(`tmp/xlsx/${name}.resaved.json`, JSON.stringify(expected(read.workbook)))
  }
})

/** Cells as Excel's check reads them back (1-based, the first 200 rows and 30 columns of each sheet's used range). */
function expected(wb: Workbook) {
  const out: { sheet: string; row: number; col: number; type: string; value: unknown; formula: string | null }[] = []
  for (const sheet of wb.sheets) {
    const used = sheet.cells.usedBounds()
    if (!used) continue
    const range = {
      r0: used.r0,
      c0: used.c0,
      r1: Math.min(used.r1, used.r0 + 199),
      c1: Math.min(used.c1, used.c0 + 29),
    }
    sheet.cells.forEach(
      range,
      (row, col, data) => {
        const v = data.value
        const type = typeof v === "number" ? "n" : typeof v === "boolean" ? "b" : isError(v) ? "e" : "s"
        const formula = !data.spill && data.input?.startsWith("=") ? data.input : null
        out.push({ sheet: sheet.name, row: row + 1, col: col + 1, type, value: isError(v) ? v.error : v, formula })
      },
      true,
    )
  }
  return out
}
