// ink.xlsx: a red pen stroke and a yellow highlighter stroke written by Cells, opened and saved again by Excel 365
// (made in the Office VM): Excel's own form of ink, with InkML's difference-encoded traces.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { isChart } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { traceXY } from "../src/xlsx/ink"

const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/ink.xlsx", import.meta.url)))

const inkOf = async (data: Uint8Array) => {
  const { workbook } = await openXlsx(data, "ink.xlsx")
  const o = workbook.sheets[0].objects[0]
  if (!o || isChart(o) || o.kind !== "shape" || !o.ink) throw new Error("no ink")
  return o
}

describe("ink in xlsx", () => {
  it("decodes InkML's explicit, first- and second-difference values", () => {
    expect(traceXY("0 0,'1323'1058,\"0-1587,264 2117")).toEqual([0, 0, 1323, 1058, 2646, 529, 4233, 2117])
    expect(traceXY("10 20, 30 40")).toEqual([10, 20, 30, 40])
  })

  it("opens Excel's ink: its strokes, colours, widths and the highlighter", async () => {
    const o = await inkOf(bytes)
    expect(o.name).toBe("Ink 1")
    expect(o.geometry).toBe("ink")
    const [pen, marker] = o.ink!.strokes
    expect(pen.color).toBe("#E71224")
    expect(pen.width).toBeCloseTo(1.89, 2)
    expect(pen.highlighter).toBeUndefined()
    const round = (xs: number[]) => xs.map((v) => Math.round(v * 1000) / 1000)
    expect(round(pen.points)).toEqual([0, 0, 0.25, 0.4, 0.5, 0.2, 0.8, 0.8])
    expect(marker).toMatchObject({ color: "#FFFF00", highlighter: true })
    expect(marker.width).toBeCloseTo(22.68, 1)
    expect(round(marker.points)).toEqual([0, 1, 1, 1])
  })

  it("saves ink as Excel does: a content part and its InkML, and reads it back the same", async () => {
    const { workbook } = await openXlsx(bytes, "ink.xlsx")
    const saved = await saveXlsx(workbook)
    const entries = await readZip(saved)
    const names = entries.map((e) => e.name)
    const text = (re: RegExp) => new TextDecoder().decode(entries.find((e) => re.test(e.name))!.data)
    const drawing = text(/drawings\/drawing\d+\.xml$/)
    expect(drawing).toMatch(/<xdr:contentPart xmlns:r="[^"]+" r:id="rId\d+"><xdr14:nvContentPartPr/)
    expect(drawing).not.toContain("AlternateContent")
    const rels = text(/drawings\/_rels\//)
    expect(rels).toContain("relationships/customXml")
    expect(names.filter((n) => /ink\/ink\d+\.xml$/.test(n))).toHaveLength(1)
    const again = await inkOf(saved)
    const first = await inkOf(bytes)
    expect(again.anchor).toEqual(first.anchor)
    expect(again.ink!.strokes.map((s) => s.points.map((v) => Math.round(v * 100)))).toEqual(
      first.ink!.strokes.map((s) => s.points.map((v) => Math.round(v * 100))),
    )
  })
})
