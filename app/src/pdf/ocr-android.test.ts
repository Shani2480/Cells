import { readFileSync } from "node:fs"
import { join } from "node:path"
import { pdfTables } from "@cells/io"
import { afterEach, describe, expect, it } from "vitest"
import type { OcrLine } from "./ocr"
import { OCR_DPI, pageAngle, runsFromOcr } from "./ocr"
import { mlkitRecognizer, recognizeJpeg } from "./ocr-android"

type Host = { __cellsOcr?: (id: string, json: string) => void; CellsOcr?: unknown }
const host = globalThis as Host

afterEach(() => {
  delete host.__cellsOcr
  delete host.CellsOcr
})

describe("text recognition through Android's bridge", () => {
  it("hands a page over and gets its lines back by id, several at once", async () => {
    const asked: string[] = []
    const bridge = { available: () => true, recognize: (id: string) => void asked.push(id) }
    const first = recognizeJpeg(bridge, "AAAA")
    const second = recognizeJpeg(bridge, "BBBB")
    const line = {
      bbox: { x0: 1, y0: 2, x1: 30, y1: 12 },
      words: [{ text: "Total", confidence: 97, bbox: { x0: 1, y0: 2, x1: 30, y1: 12 } }],
    }
    host.__cellsOcr?.(asked[1], JSON.stringify({ lines: [] }))
    host.__cellsOcr?.(asked[0], JSON.stringify({ lines: [line] }))
    expect(await first).toEqual([line])
    expect(await second).toEqual([])
  })

  it("hands the PDF over once and has Android draw each page; a page it draws blank is drawn here instead", async () => {
    const opened: number[] = []
    const asked: [number, number][] = []
    const word = { text: "Total", confidence: 97, bbox: { x0: 100, y0: 100, x1: 300, y1: 140 } }
    host.CellsOcr = {
      available: () => true,
      recognize: () => undefined,
      openPdf: (b64: string) => (opened.push(b64.length), true),
      recognizePdfPage: (id: string, index: number, scale: number) => {
        asked.push([index, scale])
        const lines = index === 0 ? [{ bbox: word.bbox, words: [word] }] : []
        queueMicrotask(() => host.__cellsOcr?.(id, JSON.stringify({ lines })))
      },
      closePdf: () => undefined,
    }
    const bytes = new Uint8Array([37, 80, 68, 70])
    const page = { width: 612, height: 792, runs: [], rules: [], scanned: true }
    let drawn = 0
    const draw = async () => {
      drawn++
      throw new Error("drawn here")
    }
    const first = await mlkitRecognizer(page, draw, { bytes, index: 0 })
    expect(first.runs.map((r) => r.text)).toEqual(["Total"])
    expect(drawn).toBe(0)
    await expect(mlkitRecognizer(page, draw, { bytes, index: 1 })).rejects.toThrow("drawn here")
    expect(drawn).toBe(1)
    expect(opened).toHaveLength(1)
    expect(asked).toEqual([
      [0, 300 / 72],
      [1, 300 / 72],
    ])
  })

  it("passes on what went wrong", async () => {
    let id = ""
    const read = recognizeJpeg({ available: () => true, recognize: (i) => void (id = i) }, "AAAA")
    host.__cellsOcr?.(id, JSON.stringify({ error: "The page couldn't be drawn for recognition." }))
    await expect(read).rejects.toThrow("couldn't be drawn")
  })
})

describe("tables in what ML Kit read on a OnePlus Pad", () => {
  // ML Kit's answers for the two demo scans (tilted 0.4°), as AndroidOcr.kt sent them.
  const scans = JSON.parse(
    readFileSync(join(__dirname, "../../../packages/io/test/fixtures/pdf/mlkit-scans.json"), "utf8"),
  ) as Record<string, { lines: OcrLine[] }>
  const tables = (name: string) =>
    pdfTables([{ width: 618, height: 850, runs: runsFromOcr(scans[name].lines, OCR_DPI / 72), rules: [] }])

  it("straightens the page by its lines' angle, so a narrow column gap lines up", () => {
    expect(pageAngle(scans["scan-statement.jpg"].lines)).toBeCloseTo(0.55, 1)
    const [t] = tables("scan-statement.jpg")
    expect(t.rows).toHaveLength(42)
    expect(t.rows[0]).toEqual(["Date", "Description", "Amount", "Balance"])
    expect(t.rows[1]).toEqual(["01/01/2026", "Coffee shop", "-5.00", "995.00"])
    expect(t.rows[41]).toEqual(["10/02/2026", "Coffee shop", "-45.00", "8,911.00"])
  })

  it("keeps a wrapped description in its row and marks what ML Kit wasn't sure of", () => {
    const [t] = tables("scan-invoice.jpg")
    expect(t.rows.map((r) => r[0])).toEqual(["Item", "A-100", "A-210", "B-007", "C-300"])
    expect(t.rows[1][1]).toBe("Logo design, three concepts and two rounds of changes to the chosen\none")
    expect(t.rows[4]).toEqual(["C-300", "Printing", "2", "€ 18.50", "€37.00"])
    expect(t.unsure?.length).toBeGreaterThan(0)
  })
})
