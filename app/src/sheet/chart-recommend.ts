/**
 * Insert › Recommended Charts: the chart types that suit the data, best first, as Excel picks them. The data's
 * shape decides: x-y pairs suit a scatter, a long or dated run a line, one small series of parts a pie, series of
 * very different sizes a combo with a second axis.
 */
export type ChartData = {
  categories: string[]
  series: { values: number[] }[]
  /** The first column holds numbers (x values), not labels. */
  numericX?: boolean
}

/** What Excel offers when there's no data to go by. */
export const DEFAULT_RECOMMENDED = ["column", "line", "pie", "hbar", "area", "scatter", "column-stacked", "doughnut"]

/** At most this many recommendations, as Excel's list. */
const MOST = 8

const DATE_LIKE =
  /^(\d{4}|q[1-4]|(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*|\d{1,4}[-/.]\d{1,2}([-/.]\d{1,4})?)$/i

export function recommendCharts(data: ChartData | null): string[] {
  if (!data || !data.series.length || !data.categories.length) return DEFAULT_RECOMMENDED
  const values = data.series.flatMap((s) => s.values.filter(Number.isFinite))
  const positive = values.every((v) => v >= 0)
  const many = data.categories.length > 12
  const dated = data.categories.filter((c) => DATE_LIKE.test(c.trim())).length >= data.categories.length * 0.8
  const out: string[] = []
  const add = (...ids: string[]) => {
    for (const id of ids) if (!out.includes(id)) out.push(id)
  }
  if (data.numericX) add("scatter", "scatter-lines-markers", "scatter-smooth-markers")
  if (data.series.length === 1) {
    if (many || (dated && data.categories.length > 6)) add("line", "column", "area", "line-markers")
    else {
      add("column", "hbar")
      if (positive && data.categories.length <= 7) add("pie", "doughnut")
      add("line", "area")
    }
    if (positive && !many) add("funnel")
  } else {
    // Series an order of magnitude apart read best as columns and a line on its own axis.
    const peaks = data.series.map((s) => Math.max(...s.values.map(Math.abs).filter(Number.isFinite), 0))
    const spread = Math.max(...peaks) / Math.max(Math.min(...peaks.filter((p) => p > 0)), 1e-9)
    if (data.series.length === 2 && spread >= 10) add("combo-column-line-secondary")
    if (many || (dated && data.categories.length > 6)) add("line", "line-markers", "column", "area", "column-stacked")
    else add("column", "hbar", "column-stacked", "line", "hbar-stacked")
    if (positive) add("area-stacked", "column-percent")
    add("combo-column-line")
  }
  add(...DEFAULT_RECOMMENDED)
  return out.slice(0, MOST)
}

/** Excel's line under a recommended chart: what the type is good for. */
export const CHART_ADVICE: Record<string, string> = {
  column:
    "A clustered column chart is used to compare values across a few categories. Use it when the order of categories is not important.",
  "column-stacked":
    "A stacked column chart is used to compare parts of a whole. Use it to show how a total is made up across categories.",
  "column-percent": "A 100% stacked column chart compares the percentage each value adds to a total across categories.",
  hbar: "A clustered bar chart is used to compare values across a few categories. Use it when the chart shows durations or when the category text is long.",
  "hbar-stacked": "A stacked bar chart is used to compare parts of a whole across categories with long names.",
  line: "A line chart is used to show trends over time (years, months and days) or categories when the order is important.",
  "line-markers":
    "A line chart with markers shows trends over time or ordered categories, with each data point marked. Use it when there are only a few points.",
  pie: "A pie chart is used to show proportions of a whole. Use it when the total of your numbers is 100%.",
  doughnut: "A doughnut chart shows proportions of a whole, like a pie, with room in the middle.",
  area: "An area chart is used to show trends over time or categories, with the area under the line filled.",
  "area-stacked":
    "A stacked area chart shows how parts of a whole change over time. Use it to stress the size of the total.",
  scatter: "A scatter chart is used to compare at least two sets of values or pairs of data, to show how they relate.",
  "scatter-lines-markers":
    "A scatter chart with straight lines and markers shows how pairs of values relate, joined in order.",
  "scatter-smooth-markers": "A scatter chart with smooth lines and markers shows a trend through pairs of values.",
  funnel: "A funnel chart shows values across the stages of a process, such as a sales pipeline.",
  "combo-column-line":
    "A clustered column and line chart is used to highlight different kinds of information, such as a total against its trend.",
  "combo-column-line-secondary":
    "A clustered column and line chart on a secondary axis is used when the values of the series are of very different sizes.",
}
