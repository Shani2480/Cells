import { normalizeChartSpec, type ChartSpec } from "./spec"
import { child, children, tree, type ChartXmlRefs, type Element } from "./xml"

/**
 * Charts in a spreadsheet: the chart reads its numbers from ranges of the workbook (Excel's c:f) rather than from
 * an embedded workbook, as Word's do.
 */

/** The groups readChartSpace reads, in its order (the series follow the groups). */
const GROUPS =
  /^(bar|bar3D|line|line3D|area|area3D|pie|pie3D|ofPie|doughnut|scatter|bubble|radar|stock|surface|surface3D)Chart$/

/** The formula of a data source (strRef, numRef, multiLvlStrRef), or null when the data is typed in. */
function formula(source: Element | undefined): string | null {
  if (!source) return null
  for (const kind of ["strRef", "numRef", "multiLvlStrRef"]) {
    const f = child(child(source, kind), "f")?.text.trim()
    if (f) return f
  }
  return null
}

/** The ranges a chart part reads, series in the order readChartSpace gives them; null when it reads none. */
export function readChartRefs(xml: string): ChartXmlRefs | null {
  const plot = child(child(tree(xml) ?? undefined, "chart"), "plotArea")
  if (!plot) return null
  const series = plot.children.filter((g) => GROUPS.test(g.name)).flatMap((g) => children(g, "ser"))
  let categories: string | null = null
  const out: ChartXmlRefs["series"] = []
  for (const s of series) {
    const values = formula(child(s, "val") ?? child(s, "yVal"))
    // A series typed into the chart: the whole chart is kept as its numbers, not its ranges.
    if (!values) return null
    categories ??= formula(child(s, "cat") ?? child(s, "xVal"))
    out.push({ name: formula(child(s, "tx")), values })
  }
  return out.length ? { categories, series: out } : null
}

/** A chart's data read from its ranges (see @cells/core's readChartData). */
export type SheetChartData = {
  categories: string[] | null
  series: { name: string | null; values: number[]; texts: string[] }[]
}

/** A chart's spec with the categories and series read from the sheet now; its look stays its own. */
export function withSheetData(look: Partial<ChartSpec>, data: SheetChartData): ChartSpec {
  const spec = normalizeChartSpec(look)
  const length = Math.max(0, ...data.series.map((s) => s.values.length))
  spec.categories = data.categories ?? Array.from({ length }, (_, i) => String(i + 1))
  spec.series = data.series.map((s, i) => ({
    ...spec.series[i],
    name: s.name ?? `Series${i + 1}`,
    values: s.values,
    ...(s.texts.some(Boolean) ? { texts: s.texts } : {}),
  }))
  if (!spec.series.length) spec.series = [{ name: "Series1", values: [] }]
  return spec
}
