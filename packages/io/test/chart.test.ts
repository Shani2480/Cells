// Charts in xlsx: a chart Excel 365 saved (fixtures/excel/features.xlsx) read with its ranges, charts Cells makes
// written as a drawing and chart parts Excel opens, and an untouched chart written back as Excel wrote it.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"
import { samples } from "./samples"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/features.xlsx", import.meta.url)))
const files = async (bytes: Uint8Array) =>
  new Map((await readZip(bytes)).map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))

describe("charts in xlsx", () => {
  it("reads Excel's chart: where it sits, its type and the ranges it reads", async () => {
    const { workbook } = await openXlsx(fixture(), "features.xlsx")
    const [chart] = workbook.sheets[0].charts
    expect(chart.name).toBe("Chart 1")
    expect(chart.anchor.from).toEqual({ col: 6, x: 14, row: 1, y: 63500 / 9525 })
    expect(chart.anchor.to.col).toBe(12)
    expect(chart.refs).toEqual({
      categories: "Sales!$A$2:$A$6",
      series: [
        { name: "Sales!$B$1", values: "Sales!$B$2:$B$6" },
        { name: "Sales!$C$1", values: "Sales!$C$2:$C$6" },
      ],
    })
    expect(chart.spec.type).toBe("column")
    // The numbers come from the cells, not the part's cache.
    expect((chart.spec.series as { values: number[] }[])[0].values).toEqual([])
  })

  it("writes an untouched chart back as Excel wrote it, and a changed one anew", async () => {
    const read = await openXlsx(fixture(), "features.xlsx")
    const before = await files(fixture())
    const same = await files(await saveXlsx(read.workbook, { source: read.source }))
    expect(same.get("xl/charts/chart1.xml")).toBe(before.get("xl/charts/chart1.xml"))
    expect(same.get("xl/worksheets/sheet1.xml")).toMatch(/<drawing r:id="rId\d+"\/>/)
    // A changed number: the part is written again with it.
    const sheet = read.workbook.sheets[0]
    sheet.cells.set(1, 1, { value: 999 })
    const changed = await files(await saveXlsx(read.workbook, { source: read.source }))
    const xml = changed.get("xl/charts/chart1.xml")!
    expect(xml).not.toBe(before.get("xl/charts/chart1.xml"))
    expect(xml).toContain("<c:f>Sales!$B$2:$B$6</c:f>")
    expect(xml).toContain("<c:v>999</c:v>")
    // Deleted: the drawing and the chart go.
    sheet.charts = []
    const gone = await files(await saveXlsx(read.workbook, { source: read.source }))
    expect([...gone.keys()].filter((n) => /chart|drawings\/drawing/.test(n))).toEqual([])
    expect(gone.get("xl/worksheets/sheet1.xml")).not.toContain("<drawing ")
    expect(gone.get("[Content_Types].xml")).not.toContain("chart")
  })

  it("writes charts Cells made and reads them back the same", async () => {
    const wb = samples.charts()
    const saved = await saveXlsx(wb)
    const out = await files(saved)
    expect(out.get("xl/worksheets/sheet1.xml")).toMatch(/<drawing r:id="rId\d+"\/>/)
    const drawing = out.get("xl/drawings/drawing1.xml")!
    expect(drawing.match(/<xdr:twoCellAnchor/g)).toHaveLength(2)
    expect(drawing).toContain("<xdr:col>4</xdr:col><xdr:colOff>95250</xdr:colOff>")
    const column = out.get("xl/charts/chart1.xml")!
    expect(column).toContain("<c:f>Sales!$A$2:$A$5</c:f>")
    expect(column).toContain("<c:f>Sales!$C$1</c:f>")
    expect(column).toContain("<c:v>170</c:v>")
    expect(column).not.toContain("externalData")
    expect(out.get("xl/charts/chart2.xml")).toContain("<c:pieChart>")
    expect(out.get("[Content_Types].xml")).toContain("drawingml.chart+xml")
    const again = await openXlsx(saved, "charts.xlsx")
    const charts = again.workbook.sheets[0].charts
    expect(charts.map((c) => c.refs)).toEqual(wb.sheets[0].charts.map((c) => c.refs))
    expect(charts.map((c) => c.anchor)).toEqual(wb.sheets[0].charts.map((c) => c.anchor))
    expect(charts.map((c) => [c.name, c.spec.type, (c.spec.title as { text: string }).text])).toEqual([
      ["Chart 1", "column", "Sales"],
      ["Chart 2", "pie", "North"],
    ])
  })
})
