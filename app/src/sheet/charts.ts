import {
  copyChart,
  chartRefsBox,
  chartRefsFor,
  chartRefRange,
  chartRefText,
  chartSeriesInRows,
  currentRegion,
  nextChartId,
  nextChartName,
  readChartData,
  switchChartRefs,
  type ChartAnchor,
  type Sheet,
  type SheetChart,
} from "@cells/core"
import {
  defaultChartSpec,
  normalizeChartSpec,
  variantOf,
  withSheetData,
  withVariant,
  type ChartSpec,
} from "@cells/chart"
import type { SheetController } from "./controller"

/** Excel's size for a new chart: 5 by 3 inches (480 by 288 pixels at 100%). */
export const CHART_WIDTH = 480
export const CHART_HEIGHT = 288

/** A chart's box on the sheet, in pixels at 100% from the sheet's top-left corner. */
export type ChartBox = { x: number; y: number; w: number; h: number }

export function chartBox(sheet: Sheet, anchor: ChartAnchor): ChartBox {
  const x = sheet.cols.offset(anchor.from.col) + anchor.from.x
  const y = sheet.rows.offset(anchor.from.row) + anchor.from.y
  const x1 = sheet.cols.offset(anchor.to.col) + anchor.to.x
  const y1 = sheet.rows.offset(anchor.to.row) + anchor.to.y
  return { x, y, w: Math.max(1, x1 - x), h: Math.max(1, y1 - y) }
}

/** The cells (and offsets into them) a box's corners fall in. */
export function anchorOf(sheet: Sheet, box: ChartBox): ChartAnchor {
  const corner = (px: number, py: number) => {
    const col = sheet.cols.indexAt(Math.max(0, px))
    const row = sheet.rows.indexAt(Math.max(0, py))
    return { row, col, x: Math.max(0, px) - sheet.cols.offset(col), y: Math.max(0, py) - sheet.rows.offset(row) }
  }
  return { from: corner(box.x, box.y), to: corner(box.x + box.w, box.y + box.h) }
}

/**
 * Charts: Insert › Charts (and Alt+F1), selecting a chart and moving, sizing and deleting it, and the Chart Design
 * tab's changes. A chart's look is its spec; its numbers come from its ranges each time it is drawn.
 */
export class ChartCommands {
  /** The selected chart's id (Excel selects a chart instead of cells; the Chart Design tab shows). */
  selected: number | null = null

  constructor(private readonly c: SheetController) {}

  /** The selected chart on the sheet shown. */
  active(): SheetChart | null {
    if (this.selected === null) return null
    return this.c.sheet.charts.find((ch) => ch.id === this.selected) ?? null
  }

  select(id: number | null) {
    if (this.selected === id) return
    if (id !== null && this.c.edit) this.c.commitEdit()
    this.selected = id
    this.c.changed()
  }

  /** The spec a chart is drawn with: its own look, with the categories and series read from its ranges now. */
  liveSpec(chart: SheetChart): ChartSpec {
    const look = chart.spec as Partial<ChartSpec>
    if (!chart.refs) return normalizeChartSpec(look)
    return withSheetData(look, readChartData(this.c.wb.sheets, this.c.wb.styles, chart.refs))
  }

  private commit(charts: SheetChart[], label: string): boolean {
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    return (
      this.c.transact([{ kind: "charts", sheet: this.c.sheetIndex, charts }], { label, before: sel, after: sel }) !==
      null
    )
  }

  private replace(chart: SheetChart, next: SheetChart | null, label: string) {
    const charts = this.c.sheet.charts.map((x) => (x.id === chart.id ? next : x)).filter((x): x is SheetChart => !!x)
    return this.commit(charts, label)
  }

  // ── Insert ──

  /**
   * Insert › a chart type (Alt+F1 for the default): a chart of the data around the selection, in the middle of the
   * window. With no data, an empty chart Select Data can fill later.
   */
  insert(variant = "column"): SheetChart | null {
    if (this.c.edit) this.c.commitEdit()
    const sheet = this.c.sheet
    const sel = this.c.selection.ranges[0]
    const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
    const range = single ? currentRegion(sheet, this.c.selection.active) : { ...sel }
    const empty = single && sheet.cells.get(range.r0, range.c0)?.value === undefined && range.r0 === range.r1
    const refs = empty ? null : chartRefsFor(sheet, range)
    const base = withVariant(defaultChartSpec("column"), variant)
    const spec: ChartSpec = { ...base, categories: [], series: [] }
    // Excel titles a one-series chart with the series' name, others "Chart Title".
    const data = refs ? readChartData(this.c.wb.sheets, this.c.wb.styles, refs) : null
    spec.title = {
      show: true,
      text: data?.series.length === 1 && data.series[0].name ? data.series[0].name : "Chart Title",
    }
    // One series: no legend (its name is the title); a pie keeps one, its slices being the categories.
    if (data?.series.length === 1 && spec.type !== "pie" && spec.type !== "doughnut") spec.legend = "none"
    const zoom = this.c.zoom || 1
    const x = Math.max(0, this.c.scrollX + (this.c.viewWidth / zoom - CHART_WIDTH) / 2)
    const y = Math.max(0, this.c.scrollY + (this.c.viewHeight / zoom - CHART_HEIGHT) / 2)
    const chart: SheetChart = {
      id: nextChartId(this.c.wb.sheets),
      name: nextChartName(sheet),
      anchor: anchorOf(sheet, { x, y, w: CHART_WIDTH, h: CHART_HEIGHT }),
      spec: spec as unknown as Record<string, unknown>,
      refs,
    }
    if (!this.commit([...sheet.charts, chart], "Insert Chart")) return null
    this.select(chart.id)
    this.c.shell.showTab?.("TabChartDesign")
    return chart
  }

  // ── Selected chart ──

  /** Moves or sizes a chart to a box (pixels at 100%). */
  place(id: number, box: ChartBox) {
    const chart = this.c.sheet.charts.find((x) => x.id === id)
    if (!chart) return
    const next = copyChart(chart)
    next.anchor = anchorOf(this.c.sheet, box)
    this.replace(chart, next, "Move Chart")
  }

  delete() {
    const chart = this.active()
    if (!chart) return
    this.replace(chart, null, "Delete Chart")
    this.select(null)
  }

  /** A change to the selected chart's look (Chart Design and the buttons beside it), as one undo step. */
  change(change: (spec: ChartSpec) => ChartSpec, label = "Chart Format") {
    const chart = this.active()
    if (!chart) return
    const live = this.liveSpec(chart)
    const next = copyChart(chart)
    next.spec = keepLook(change(live), Boolean(chart.refs)) as unknown as Record<string, unknown>
    this.replace(chart, next, label)
  }

  /** The selected chart's spec as drawn (null when no chart is selected). */
  spec(): ChartSpec | null {
    const chart = this.active()
    return chart ? this.liveSpec(chart) : null
  }

  variant(): string | null {
    const spec = this.spec()
    return spec ? variantOf(spec) : null
  }

  /** Change Chart Type. */
  setVariant(id: string) {
    this.change((spec) => withVariant(spec, id), "Change Chart Type")
  }

  /** Select Data: the chart's data range as the dialog shows it ("=Sheet1!$A$1:$C$4"), or "" when it has none. */
  dataRange(chart = this.active()): string {
    if (!chart?.refs) return ""
    const box = chartRefsBox(this.c.wb.sheets, chart.refs)
    return box ? "=" + chartRefText(box.sheet.name, box.range) : ""
  }

  /**
   * Select Data › Chart data range: the series read again from a new range (series kept running the same way), or
   * false when the text isn't a range. `byRow` switches them.
   */
  setData(text: string, byRow?: boolean): boolean {
    const chart = this.active()
    if (!chart) return false
    const at = chartRefRange(this.c.wb.sheets, text)
    if (!at) return false
    const rows = byRow ?? (chart.refs ? chartSeriesInRows(this.c.wb.sheets, chart.refs) : undefined)
    const next = copyChart(chart)
    next.refs = chartRefsFor(at.sheet, at.range, rows)
    this.replace(chart, next, "Select Data")
    return true
  }

  /** Select Data › Remove: a series out of the chart. */
  removeSeries(index: number) {
    const chart = this.active()
    if (!chart?.refs || chart.refs.series.length <= 1) return
    const next = copyChart(chart)
    next.refs!.series.splice(index, 1)
    const spec = normalizeChartSpec(chart.spec as Partial<ChartSpec>)
    spec.series.splice(index, 1)
    next.spec = spec as unknown as Record<string, unknown>
    this.replace(chart, next, "Select Data")
  }

  /** Switch Row/Column: the same data with the series taken the other way. */
  switchRowColumn() {
    const chart = this.active()
    if (!chart?.refs) return
    const refs = switchChartRefs(this.c.wb.sheets, chart.refs)
    if (!refs) return
    const next = copyChart(chart)
    next.refs = refs
    // Series formats belong to the old series.
    const spec = normalizeChartSpec(chart.spec as Partial<ChartSpec>)
    spec.series = []
    next.spec = spec as unknown as Record<string, unknown>
    this.replace(chart, next, "Switch Row/Column")
  }
}

/**
 * What a chart keeps of a spec: everything but the numbers when they come from cells (they are read again each time),
 * so a saved spec never carries stale data.
 */
function keepLook(spec: ChartSpec, fromCells: boolean): ChartSpec {
  if (!fromCells) return spec
  return {
    ...spec,
    categories: [],
    series: spec.series.map((s) => ({ ...s, values: [], texts: undefined })),
  }
}
