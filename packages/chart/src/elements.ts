// From Inkspire (src/components/chart-menus.tsx): the chart elements Add Chart Element and the Chart Elements
// button offer, and thumbnails' text drawn as bars. Pure: no editor, no React.
import {
  chartHasAxes,
  isStatisticChart,
  type ChartSpec,
  type LabelPosition,
  type LegendPosition,
  type TrendlineKind,
} from "./spec"

const attr = (attrs: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(attrs)?.[1]

/**
 * A small thumbnail's text as short bars, as Word draws its gallery thumbnails: words a few pixels high are only a
 * smudge, but a bar still shows where the title, legend and labels are.
 */
export function textAsBars(svg: string) {
  return svg.replace(/<text([^>]*)>([^<]*)<\/text>/g, (_match, attrs: string, content: string) => {
    const size = Number(attr(attrs, "font-size")) || 12
    const length = content.replace(/&[a-z#0-9]+;/gi, "x").trim().length
    if (!length) {
      return ""
    }
    const width = Math.max(size * 0.8, length * size * 0.5)
    const x = Number(attr(attrs, "x") ?? 0)
    const y = Number(attr(attrs, "y") ?? 0)
    const anchor = attr(attrs, "text-anchor")
    const left = anchor === "middle" ? x - width / 2 : anchor === "end" ? x - width : x
    const transform = attr(attrs, "transform")
    const part = attr(attrs, "data-chart-part")
    return (
      `<rect${part ? ` data-chart-part="${part}"` : ""} x="${left.toFixed(1)}" y="${(y - size * 0.68).toFixed(1)}" width="${width.toFixed(1)}" height="${(size * 0.55).toFixed(1)}" rx="${(size * 0.15).toFixed(1)}"` +
      ` fill="${attr(attrs, "fill") ?? "#595959"}" fill-opacity="0.55"${transform ? ` transform="${transform}"` : ""}/>`
    )
  })
}

// --- Add Chart Element ------------------------------------------------------------------------------------------

export type Choice = { label: string; change: (spec: ChartSpec) => ChartSpec; on: (spec: ChartSpec) => boolean }

const LABEL_CHOICES: [LabelPosition, string][] = [
  ["none", "None"],
  ["center", "Center"],
  ["insideEnd", "Inside End"],
  ["insideBase", "Inside Base"],
  ["outsideEnd", "Outside End"],
]

const LEGEND_CHOICES: [LegendPosition, string][] = [
  ["none", "None"],
  ["right", "Right"],
  ["top", "Top"],
  ["left", "Left"],
  ["bottom", "Bottom"],
]

const TRENDLINE_CHOICES: [TrendlineKind, string][] = [
  ["none", "None"],
  ["linear", "Linear"],
  ["exponential", "Exponential"],
  ["logarithmic", "Logarithmic"],
  ["movingAverage", "Two Period Moving Average"],
]

const ERROR_BAR_CHOICES: [ChartSpec["errorBars"]["kind"], string, number?][] = [
  ["none", "None"],
  ["stderr", "Standard Error"],
  ["percent", "Percentage", 5],
  ["stddev", "Standard Deviation", 1],
  ["fixed", "Fixed Value", 1],
]

/** A chart with no axes: pie and doughnut, and a map. */
function isRound(spec: ChartSpec) {
  return spec.type === "pie" || spec.type === "doughnut" || spec.type === "map"
}

/** Word's chart elements, each with its placements and the Format pane part its More Options opens. */
/** Word lists the gridlines by direction, whichever axis draws them (a bar chart's value axis runs across). */
const GRIDLINE_ORDER = [
  "Primary Major Horizontal",
  "Primary Major Vertical",
  "Primary Minor Horizontal",
  "Primary Minor Vertical",
]

function sortedGridlines(choices: Choice[]) {
  return [...choices].sort((a, b) => GRIDLINE_ORDER.indexOf(a.label) - GRIDLINE_ORDER.indexOf(b.label))
}

export function chartElements(spec: ChartSpec): {
  id: string
  label: string
  part: string
  more: string
  shown: boolean
  toggle: (spec: ChartSpec) => ChartSpec
  choices: Choice[]
  disabled?: boolean
}[] {
  // No axes: no axis titles or gridlines; Word's 2016 kinds take none of the extras drawn from the series.
  const round = !chartHasAxes(spec.type)
  const ownRules = isStatisticChart(spec.type)
  const horizontal = spec.type === "bar"
  // Word's funnel has one axis, its categories down the side: no value axis, axis titles or gridlines.
  const funnel = spec.type === "funnel"
  // A chart with a second value axis offers it as Secondary Vertical (a Pareto's percentages, a combo's).
  const secondary = Boolean(spec.pareto) || spec.series.some((serie) => serie.secondary)
  const trendable = !round && !ownRules && spec.type !== "radar" && spec.grouping === "clustered"
  return [
    {
      id: "axes",
      label: "Axes",
      part: "axis:value",
      more: "More Axis Options…",
      shown: !round && (spec.axes.category || (!funnel && spec.axes.value)),
      disabled: round,
      toggle: (current) => {
        const on = !(current.axes.category || (!funnel && current.axes.value))
        return { ...current, axes: { category: on, value: on && !funnel, secondary: on } }
      },
      choices: funnel
        ? [
            {
              label: "Primary Vertical",
              on: (current) => current.axes.category,
              change: (current) => ({ ...current, axes: { ...current.axes, category: !current.axes.category } }),
            },
          ]
        : [
            {
              label: "Primary Horizontal",
              on: (current) => (horizontal ? current.axes.value : current.axes.category),
              change: (current) =>
                horizontal
                  ? { ...current, axes: { ...current.axes, value: !current.axes.value } }
                  : { ...current, axes: { ...current.axes, category: !current.axes.category } },
            },
            {
              label: "Primary Vertical",
              on: (current) => (horizontal ? current.axes.category : current.axes.value),
              change: (current) =>
                horizontal
                  ? { ...current, axes: { ...current.axes, category: !current.axes.category } }
                  : { ...current, axes: { ...current.axes, value: !current.axes.value } },
            },
            ...(secondary
              ? [
                  {
                    label: "Secondary Vertical",
                    on: (current: ChartSpec) => current.axes.secondary !== false,
                    change: (current: ChartSpec) => ({
                      ...current,
                      axes: { ...current.axes, secondary: current.axes.secondary === false },
                    }),
                  },
                ]
              : []),
          ],
    },
    {
      id: "axisTitles",
      label: "Axis Titles",
      part: "axisTitle:value",
      more: "More Axis Title Options…",
      shown: !round && !funnel && (spec.axisTitles.category.show || spec.axisTitles.value.show),
      disabled: round || funnel,
      toggle: (current) => {
        const on = !(current.axisTitles.category.show || current.axisTitles.value.show)
        return {
          ...current,
          axisTitles: {
            category: { ...current.axisTitles.category, show: on },
            value: { ...current.axisTitles.value, show: on },
          },
        }
      },
      choices: (["category", "value"] as const).map((axis) => ({
        label: (axis === "category") !== horizontal ? "Primary Horizontal" : "Primary Vertical",
        on: (current: ChartSpec) => current.axisTitles[axis].show,
        change: (current: ChartSpec) => ({
          ...current,
          axisTitles: {
            ...current.axisTitles,
            [axis]: { ...current.axisTitles[axis], show: !current.axisTitles[axis].show },
          },
        }),
      })),
    },
    {
      id: "title",
      label: "Chart Title",
      part: "title",
      more: "More Title Options…",
      shown: spec.title.show,
      toggle: (current) => ({ ...current, title: { ...current.title, show: !current.title.show } }),
      choices: [
        {
          label: "None",
          on: (current) => !current.title.show,
          change: (current) => ({ ...current, title: { ...current.title, show: false } }),
        },
        {
          label: "Above Chart",
          on: (current) => current.title.show,
          change: (current) => ({ ...current, title: { ...current.title, show: true } }),
        },
      ],
    },
    {
      id: "dataLabels",
      label: "Data Labels",
      part: "dataLabels",
      more: "More Data Label Options…",
      shown: spec.dataLabels !== "none",
      toggle: (current) => ({
        ...current,
        dataLabels: current.dataLabels === "none" ? (isRound(current) ? "bestFit" : "outsideEnd") : "none",
      }),
      choices: (round
        ? ([
            ["none", "None"],
            ["center", "Center"],
            ["insideEnd", "Inside End"],
            ["outsideEnd", "Outside End"],
            ["bestFit", "Best Fit"],
          ] as [LabelPosition, string][])
        : spec.type === "line" || spec.type === "scatter" || spec.type === "area" || spec.type === "radar"
          ? ([
              ["none", "None"],
              ["center", "Center"],
              ["outsideEnd", "Above"],
            ] as [LabelPosition, string][])
          : LABEL_CHOICES
      ).map(([position, label]) => ({
        label,
        on: (current: ChartSpec) => current.dataLabels === position,
        change: (current: ChartSpec) => ({ ...current, dataLabels: position }),
      })),
    },
    {
      id: "dataTable",
      label: "Data Table",
      part: "dataTable",
      more: "More Data Table Options…",
      shown: spec.dataTable,
      disabled: round || ownRules || horizontal || spec.type === "scatter" || spec.type === "radar",
      toggle: (current) => ({ ...current, dataTable: !current.dataTable }),
      choices: [
        { label: "None", on: (current) => !current.dataTable, change: (current) => ({ ...current, dataTable: false }) },
        {
          label: "With Legend Keys",
          on: (current) => current.dataTable,
          change: (current) => ({ ...current, dataTable: true }),
        },
      ],
    },
    {
      id: "gridlines",
      label: "Gridlines",
      part: "gridlines",
      more: "More Gridline Options…",
      shown:
        !round &&
        !funnel &&
        (spec.gridlines.major ||
          spec.gridlines.minor ||
          spec.gridlines.vertical ||
          Boolean(spec.gridlines.minorVertical)),
      disabled: round || funnel,
      toggle: (current) => ({ ...current, gridlines: { ...current.gridlines, major: !current.gridlines.major } }),
      choices: sortedGridlines([
        {
          label: horizontal ? "Primary Major Vertical" : "Primary Major Horizontal",
          on: (current) => current.gridlines.major,
          change: (current) => ({ ...current, gridlines: { ...current.gridlines, major: !current.gridlines.major } }),
        },
        {
          label: horizontal ? "Primary Major Horizontal" : "Primary Major Vertical",
          on: (current) => current.gridlines.vertical,
          change: (current) => ({
            ...current,
            gridlines: { ...current.gridlines, vertical: !current.gridlines.vertical },
          }),
        },
        {
          label: horizontal ? "Primary Minor Vertical" : "Primary Minor Horizontal",
          on: (current) => current.gridlines.minor,
          change: (current) => ({ ...current, gridlines: { ...current.gridlines, minor: !current.gridlines.minor } }),
        },
        {
          label: horizontal ? "Primary Minor Horizontal" : "Primary Minor Vertical",
          on: (current) => Boolean(current.gridlines.minorVertical),
          change: (current) => ({
            ...current,
            gridlines: { ...current.gridlines, minorVertical: !current.gridlines.minorVertical },
          }),
        },
      ]),
    },
    {
      id: "legend",
      label: "Legend",
      part: "legend",
      more: "More Legend Options…",
      shown: spec.legend !== "none",
      toggle: (current) => ({ ...current, legend: current.legend === "none" ? "right" : "none" }),
      choices: LEGEND_CHOICES.map(([position, label]) => ({
        label,
        on: (current: ChartSpec) => current.legend === position,
        change: (current: ChartSpec) => ({ ...current, legend: position }),
      })),
    },
    {
      id: "errorBars",
      label: "Error Bars",
      part: "series:0",
      more: "More Error Bars Options…",
      shown: spec.errorBars.kind !== "none",
      disabled: round || ownRules || spec.type === "radar",
      toggle: (current) => ({
        ...current,
        errorBars: { ...current.errorBars, kind: current.errorBars.kind === "none" ? "stderr" : "none" },
      }),
      choices: ERROR_BAR_CHOICES.map(([kind, label, amount]) => ({
        label,
        on: (current: ChartSpec) => current.errorBars.kind === kind,
        change: (current: ChartSpec) => ({
          ...current,
          errorBars: { kind, amount: amount ?? current.errorBars.amount },
        }),
      })),
    },
    {
      id: "lines",
      label: "Lines",
      part: "series:0",
      more: "More Lines Options…",
      shown: spec.dropLines || spec.highLowLines,
      disabled: !(spec.type === "line" || spec.type === "area"),
      toggle: (current) => ({
        ...current,
        dropLines: !(current.dropLines || current.highLowLines),
        highLowLines: false,
      }),
      choices: [
        {
          label: "None",
          on: (current) => !current.dropLines && !current.highLowLines,
          change: (current) => ({ ...current, dropLines: false, highLowLines: false }),
        },
        {
          label: "Drop Lines",
          on: (current) => current.dropLines,
          change: (current) => ({ ...current, dropLines: !current.dropLines }),
        },
        {
          label: "High-Low Lines",
          on: (current) => current.highLowLines,
          change: (current) => ({ ...current, highLowLines: !current.highLowLines }),
        },
      ],
    },
    {
      id: "upDownBars",
      label: "Up/Down Bars",
      part: "series:0",
      more: "More Up/Down Bars Options…",
      shown: spec.upDownBars,
      disabled: !(spec.type === "line" && spec.series.length > 1),
      toggle: (current) => ({ ...current, upDownBars: !current.upDownBars }),
      choices: [
        {
          label: "None",
          on: (current) => !current.upDownBars,
          change: (current) => ({ ...current, upDownBars: false }),
        },
        {
          label: "Up/Down Bars",
          on: (current) => current.upDownBars,
          change: (current) => ({ ...current, upDownBars: true }),
        },
      ],
    },
    {
      id: "trendline",
      label: "Trendline",
      part: "series:0",
      more: "More Trendline Options…",
      shown: spec.series.some((serie) => serie.trendline && serie.trendline !== "none"),
      disabled: !trendable,
      toggle: (current) => {
        const on = !current.series.some((serie) => serie.trendline && serie.trendline !== "none")
        return {
          ...current,
          series: current.series.map((serie, index) =>
            index === 0 ? { ...serie, trendline: on ? "linear" : "none" } : { ...serie, trendline: "none" },
          ),
        }
      },
      choices: TRENDLINE_CHOICES.map(([kind, label]) => ({
        label,
        on: (current: ChartSpec) => (current.series[0]?.trendline ?? "none") === kind,
        change: (current: ChartSpec) => ({
          ...current,
          series: current.series.map((serie, index) => (index === 0 ? { ...serie, trendline: kind } : serie)),
        }),
      })),
    },
  ]
}
