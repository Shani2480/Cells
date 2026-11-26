// @vitest-environment node
import { readFileSync } from "node:fs"
import { join } from "node:path"
import { pdfTables } from "@cells/io"
import { createWorker, OEM } from "tesseract.js"
import { EN_US } from "@cells/core"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { numberFrom } from "@/sheet/power-query"
import { recognizePage, runsFromOcr, type OcrEngine } from "./ocr"

/** Scans of the demo statement's first page and the demo invoice: 300 dpi, grey, a little tilted, blurred, grainy. */
const fixture = (name: string) => readFileSync(join(__dirname, "../../../packages/io/test/fixtures/pdf", name))
const DPI = 300

let worker: Awaited<ReturnType<typeof createWorker>>
beforeAll(async () => {
  worker = await createWorker("eng", OEM.LSTM_ONLY, {
    langPath: join(__dirname, "../../../node_modules/@tesseract.js-data/eng/4.0.0_best_int"),
    cacheMethod: "none",
    gzip: true,
  })
}, 60_000)
afterAll(() => worker?.terminate())

const tables = async (name: string) => {
  const lines = await recognizePage(worker as unknown as OcrEngine, fixture(name))
  const runs = runsFromOcr(lines, DPI / 72)
  return pdfTables([{ width: 612, height: 792, runs, rules: [] }])
}

describe("From PDF on scanned pages", () => {
  it("finds a scanned statement's table, every row and column", async () => {
    const [t] = await tables("scan-statement.jpg")
    expect(t.kind).toBe("Table")
    expect(t.rows).toHaveLength(42)
    expect(t.rows[0]).toEqual(["Date", "Description", "Amount", "Balance"])
    expect(t.rows[1]).toEqual(["01/01/2026", "Coffee shop", "-5.00", "995.00"])
    expect(t.rows[3]).toEqual(["03/01/2026", "Salary", "2,500.00", "3,453.00"])
    expect(t.rows[41]).toEqual(["10/02/2026", "Coffee shop", "-45.00", "8,911.00"])
  }, 60_000)

  it("finds a scanned invoice's table, a doubtful row read again", async () => {
    const [t] = await tables("scan-invoice.jpg")
    expect(t.rows[0]).toEqual(["Item", "Description", "Qty", "Price", "Total"])
    expect(t.rows.map((r) => r[0])).toEqual(["Item", "A-100", "A-210", "B-007", "C-300"])
    expect(t.rows[1][1]).toBe("Logo design, three concepts and two rounds of changes to the chosen\none")
    expect(t.rows[3][1]).toBe("Website banner in four sizes for the spring campaign, with source files")
    // The figures as the type step reads them (recognition sets the € signs and spaces its own way).
    expect(t.rows.slice(1).map((r) => r.slice(2).map((v) => numberFrom(v, EN_US)))).toEqual([
      [1, 450, 450],
      [500, 0.12, 60],
      [4, 35, 140],
      [2, 18.5, 37],
    ])
  }, 60_000)
})
