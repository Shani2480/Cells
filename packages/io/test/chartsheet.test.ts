import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { isChart } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"

/** Excel's file (mock data): a chart sheet "Units Chart" (a column chart of Sales!A1:B5) before the sheet "Sales". */
const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/chartsheet.xlsx", import.meta.url)))

const text = (files: { name: string; data: Uint8Array }[], re: RegExp) =>
  new TextDecoder().decode(files.find((f) => re.test(f.name))!.data)

describe("chart sheets in xlsx", () => {
  it("reads a chart sheet as a sheet holding its chart, printed across the page", async () => {
    const { workbook, warnings } = await openXlsx(bytes, "chartsheet.xlsx")
    expect(warnings).toEqual([])
    const [sheet, data] = workbook.sheets
    expect(sheet).toMatchObject({ name: "Units Chart", chartSheet: true })
    expect(sheet.objects).toHaveLength(1)
    const chart = sheet.objects[0]
    expect(isChart(chart) && chart.refs?.series[0]).toMatchObject({ values: "Sales!$B$2:$B$5" })
    expect(sheet.pageSetup.orientation).toBe("landscape")
    expect(data.chartSheet).toBe(false)
  })

  it("saves it back as a chart sheet, in the same parts, with an absolute anchor", async () => {
    const r = await openXlsx(bytes, "chartsheet.xlsx")
    const files = await readZip(await saveXlsx(r.workbook, { source: r.source, views: r.views }))
    expect(text(files, /_rels\/workbook\.xml\.rels$/)).toMatch(
      /relationships\/chartsheet" Target="chartsheets\/sheet1.xml"/,
    )
    expect(text(files, /^xl\/workbook\.xml$/)).toMatch(
      /<sheets><sheet name="Units Chart" sheetId="1" r:id="[^"]+"\/><sheet name="Sales"/,
    )
    const part = text(files, /^xl\/chartsheets\/sheet1\.xml$/)
    expect(part).toMatch(/<sheetView tabSelected="1" zoomScale="101" workbookViewId="0" zoomToFit="1"\/>/)
    expect(part).toMatch(/<pageSetup orientation="landscape"\/><drawing r:id="rId\d+"\/><\/chartsheet>$/)
    const drawing = text(files, /^xl\/drawings\/drawing\d+\.xml$/)
    expect(drawing).toMatch(
      /<xdr:absoluteAnchor[^>]*><xdr:pos x="0" y="0"\/><xdr:ext cx="\d+" cy="\d+"\/><xdr:graphicFrame/,
    )
    expect(drawing).not.toContain("<xdr:from>")
    const again = await openXlsx(await saveXlsx(r.workbook, { source: r.source }), "again.xlsx")
    expect(again.workbook.sheets[0]).toMatchObject({ chartSheet: true, name: "Units Chart" })
  })

  it("writes a chart sheet into a new file, and keeps its chart's size", async () => {
    const { workbook } = await openXlsx(bytes, "chartsheet.xlsx")
    const files = await readZip(await saveXlsx(workbook))
    const drawing = text(files, /^xl\/drawings\/drawing\d+\.xml$/)
    // Excel's own size for this chart, read and written back.
    expect(drawing).toMatch(/<xdr:ext cx="866\d{4}" cy="629\d{4}"\/>/)
    expect(files.some((f) => /^xl\/chartsheets\/sheet\d+\.xml$/.test(f.name))).toBe(true)
  })
})
