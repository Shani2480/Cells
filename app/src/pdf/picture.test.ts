import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import type { OcrLine } from "./ocr"
import { pictureTables } from "./picture"

// ML Kit's reading of the two demo scans (on a OnePlus Pad), as pictures: boxes in the photo's pixels.
const scans = JSON.parse(
  readFileSync(join(__dirname, "../../../packages/io/test/fixtures/pdf/mlkit-scans.json"), "utf8"),
) as Record<string, { lines: OcrLine[] }>

/** Lines scaled as a smaller picture would give them (a phone photo or a screenshot at another size). */
const scaled = (lines: OcrLine[], f: number): OcrLine[] =>
  lines.map((l) => ({
    ...l,
    bbox: { x0: l.bbox.x0 * f, y0: l.bbox.y0 * f, x1: l.bbox.x1 * f, y1: l.bbox.y1 * f },
    words: l.words.map((w) => ({
      ...w,
      bbox: { x0: w.bbox.x0 * f, y0: w.bbox.y0 * f, x1: w.bbox.x1 * f, y1: w.bbox.y1 * f },
    })),
  }))

describe("Data from Picture: the table in a picture", () => {
  it("takes the table, and offers everything in the picture when there's more (a title, an address)", () => {
    const read = pictureTables(scans["scan-invoice.jpg"].lines)
    expect(read.table.rows.map((r) => r[0])).toEqual(["Item", "A-100", "A-210", "B-007", "C-300"])
    expect(read.table.rows[0]).toEqual(["Item", "Description", "Qty", "Price", "Total"])
    expect(read.table.unsure.length).toBeGreaterThan(0)
    expect(read.whole?.rows[0][0]).toBe("Invoice")
    expect(read.whole?.rows.some((r) => r.join(" ").includes("Please pay within 30 days."))).toBe(true)
  })

  it("finds the same table whatever the picture's size", () => {
    for (const f of [0.35, 1, 1.6]) {
      const { table } = pictureTables(scaled(scans["scan-statement.jpg"].lines, f))
      expect(table.rows).toHaveLength(42)
      expect(table.rows[1]).toEqual(["01/01/2026", "Coffee shop", "-5.00", "995.00"])
    }
  })

  it("reads nothing from a picture without text", () => {
    expect(pictureTables([]).table.rows).toEqual([])
  })
})
