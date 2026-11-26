import { describe, expect, it } from "vitest"
import { chartSpaceXml, defaultChartSpec, readChartSpace, withVariant, type ChartSpec } from "../src"

/** A chart of three months and two series (mock data), as the Format tab would leave it. */
function chart(type = "column"): ChartSpec {
  const spec = withVariant(defaultChartSpec("column"), type)
  return {
    ...spec,
    categories: ["Jan", "Feb", "Mar"],
    series: [
      { name: "North", values: [12, 19, 15] },
      { name: "South", values: [8, 11, 14] },
    ],
    axisTitles: { category: { show: true, text: "Month" }, value: { show: true, text: "Units" } },
    legend: "right",
    dataLabels: "outsideEnd",
  }
}

describe("chart part formats (the Format tab) in chart XML", () => {
  it("keeps every part's own fill, line, width and text through a save and read", () => {
    const format: ChartSpec["format"] = {
      chartArea: { fill: "#fff2cc", line: "#c00000", lineWidth: 2 },
      plotArea: { fill: "#deebf7", line: "#2f5597", lineWidth: 1.5 },
      title: { color: "#c00000", fontSize: 18, bold: true },
      "axisTitle:category": { color: "#2f5597", fontSize: 12, bold: false },
      "axis:value": { color: "#7030a0", fontSize: 11 },
      "axis:category": { line: "#000000", lineWidth: 1, color: "#00b050", bold: true },
      gridlines: { line: "#ff0000", lineWidth: 1.25 },
      legend: { color: "#7030a0", fontSize: 12, bold: true },
      dataLabels: { color: "#ffffff", fontSize: 8 },
      "series:0": { fill: "#ffc000", line: "#000000", lineWidth: 1 },
    }
    const back = readChartSpace(chartSpaceXml({ ...chart(), format }))!
    for (const [part, own] of Object.entries(format)) expect(back.format[part], part).toMatchObject(own)
  })

  it("reads nothing extra for a chart left as its style draws it", () => {
    const back = readChartSpace(chartSpaceXml(chart()))!
    const extra = Object.keys(back.format).filter((part) => !part.startsWith("series:"))
    expect(extra).toEqual([])
    for (const [part, own] of Object.entries(back.format)) expect(Object.keys(own), part).toEqual(["fill"])
  })

  it("keeps a line chart's line colour and weight, and no line on a column chart's chart area", () => {
    const line = readChartSpace(
      chartSpaceXml({ ...chart("line"), format: { "series:1": { fill: "#ff0000", lineWidth: 4 } } }),
    )!
    expect(line.format["series:1"]).toMatchObject({ fill: "#ff0000", lineWidth: 4 })
    const plain = readChartSpace(chartSpaceXml({ ...chart(), format: { chartArea: { line: "none" } } }))!
    expect(plain.format.chartArea).toEqual({ line: "none" })
  })
})

describe("theme colours in chart XML", () => {
  it("keeps a series' theme colour as a link to the theme, and leaves Excel's automatic accents automatic", () => {
    const format: ChartSpec["format"] = {
      "series:0": { fill: "scheme:accent1" },
      "series:1": { fill: "scheme:accent5/lumMod=75000" },
      plotArea: { fill: "scheme:accent4/lumMod=20000/lumOff=80000" },
    }
    const xml = chartSpaceXml({ ...chart(), format })
    expect(xml).toContain('<a:schemeClr val="accent5"><a:lumMod val="75000"/></a:schemeClr>')
    const back = readChartSpace(xml)!
    expect(back.format["series:0"]).toBeUndefined()
    expect(back.format["series:1"]).toEqual({ fill: "scheme:accent5/lumMod=75000" })
    expect(back.format.plotArea).toMatchObject({ fill: "scheme:accent4/lumMod=20000/lumOff=80000" })
  })
})
