// Sparklines in xlsx: a file Excel 365 made (fixtures/excel/sparklines.xlsx) with a line group (markers, high, low,
// first, last, 1.5 pt), a column (negative points, axis, own colour), a win/loss, a line with every option of the
// Hidden & Empty Cells and axis dialogs, and one on a date axis.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  sparklineAt,
  sparklinePoints,
  sparklineBounds,
  Workbook,
  newSparklineGroup,
  sparkStyleIndex,
} from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/sparklines.xlsx", import.meta.url)))

describe("sparklines in xlsx", () => {
  it("reads every group and option Excel wrote", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "sparklines.xlsx")
    const groups = wb.sheets[0].sparklines
    expect(groups).toHaveLength(5)
    const at = (row: number) => sparklineAt(groups, row, 5)!.group
    const line = at(1)
    expect(line.sparklines.map((s) => s.source)).toEqual(["Sheet1!A2:E2", "Sheet1!A3:E3", "Sheet1!A4:E4"])
    expect(line).toMatchObject({
      type: "line",
      lineWeight: 1.5,
      markers: true,
      high: true,
      low: true,
      first: true,
      last: true,
    })
    expect(line.colors.series).toBe("#376092")
    expect(at(5)).toMatchObject({ type: "column", negative: true, displayXAxis: true })
    expect(at(5).colors.series).toBe("#A05020")
    expect(at(7)).toMatchObject({ type: "stacked", negative: true })
    expect(at(9)).toMatchObject({
      emptyAs: "span",
      showHidden: true,
      minAxis: "custom",
      maxAxis: "custom",
      manualMin: 0,
      manualMax: 10,
      rightToLeft: true,
    })
    expect(at(11).dateRange).toBe("Sheet1!A13:E13")
    expect(sparklinePoints(wb.sheets, at(9), at(9).sparklines[0]).values).toEqual([2, null, 5, 1, 4])
    const dated = sparklinePoints(wb.sheets, at(11), at(11).sparklines[0])
    expect(dated.x.map((x) => Math.round(x * 50))).toEqual([0, 5, 20, 21, 50])
  })

  it("writes them back as Excel does, and they read back the same", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "sparklines.xlsx")
    const saved = await saveXlsx(wb)
    const zip = await readZip(saved)
    const sheet = new TextDecoder().decode(zip.find((e) => e.normalizedName === "xl/worksheets/sheet1.xml")!.data)
    expect(sheet.match(/<x14:sparklineGroup\b/g)).toHaveLength(5)
    expect(sheet).toContain(
      `<x14:sparklineGroup lineWeight="1.5" displayEmptyCellsAs="gap" markers="1" high="1" low="1" first="1" last="1">`,
    )
    expect(sheet).toContain(`<x14:sparkline><xm:f>Sheet1!A2:E2</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline>`)
    const again = await openXlsx(saved, "x.xlsx")
    const strip = (g: unknown) => JSON.parse(JSON.stringify(g, (k, v) => (k === "uid" ? undefined : v)))
    expect(strip(again.workbook.sheets[0].sparklines)).toEqual(strip(wb.sheets[0].sparklines))
  })

  it("a new group is Excel's default style and writes theme colours", async () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    ;[1, 3, 2].forEach((v, c) => s.cells.set(0, c, { value: v }))
    s.sparklines = [newSparklineGroup("line", [{ row: 0, col: 3, source: "Sheet1!A1:C1" }])]
    expect(sparkStyleIndex(s.sparklines[0].colors)).toBe(0)
    const pts = sparklinePoints(wb.sheets, s.sparklines[0], s.sparklines[0].sparklines[0])
    expect(sparklineBounds(s.sparklines[0], [pts])).toEqual([{ min: 1, max: 3 }])
    const zip = await readZip(await saveXlsx(wb))
    const xml = new TextDecoder().decode(zip.find((e) => e.normalizedName === "xl/worksheets/sheet1.xml")!.data)
    expect(xml).toContain(`<x14:colorSeries theme="4" tint="-0.499984740745262"/>`)
  })
})
