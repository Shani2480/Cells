import { describe, expect, it } from "vitest"
import { defaultChartSpec, renderChart, withVariant } from "../src"
import { chartSpaceXml } from "../src/xml"

const line = () => {
  const spec = withVariant(defaultChartSpec("line"), "line")
  spec.categories = ["1", "2", "3", "4", "5"]
  spec.series = [{ name: "Sales", values: [100, 120, 0, 130, 140], blanks: [2] }]
  return spec
}

describe("empty cells in a chart", () => {
  it("break a line", () => {
    const svg = renderChart(line(), 400, 240)
    const path = /<path d="(M[^"]*)" fill="none" stroke="[^"]*" stroke-width="2\.25"/.exec(svg)?.[1] ?? ""
    expect(path.match(/M/g)).toHaveLength(2)
  })

  it("have no point in the chart's cached numbers", () => {
    const xml = chartSpaceXml(line())
    expect(xml).toContain(
      '<c:ptCount val="5"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>120</c:v></c:pt><c:pt idx="3">',
    )
  })
})
