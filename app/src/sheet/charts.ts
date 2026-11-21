import {
  pivotChartButtons,
  pivotChartData,
  type PivotCache,
  type PivotChartLink,
  type PivotTable,
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
  type ChartRefs,
  type Range,
  type Sheet,
  type SheetChart,
  type Theme,
  type Op,
  chartSheetSnapshot,
  nextChartSheetName,
  sheetNameProblem,
} from "@cells/core"
import {
  defaultChartSpec,
  normalizeChartSpec,
  renderChart,
  variantOf,
  withSheetData,
  withVariant,
  type ChartSpec,
} from "@cells/chart"
import type { SheetController } from "./controller"
import { chartHtml, copiedChartInfo, svgPicture, type CopiedChartInfo } from "./chart-clipboard"

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
/** A theme's six accents as charts take them. */
export function themeAccents(theme: Theme): string[] {
  return (["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const).map((s) => `#${theme.colors[s]}`)
}

export class ChartCommands {
  /**
   * The selected chart's id (Excel selects a chart instead of cells; the Chart Design tab shows): the object
   * selection when it's one chart.
   */
  get selected(): number | null {
    // A chart sheet is its chart: it's always the one selected.
    if (this.c.sheet.chartSheet) return this.c.sheet.charts[0]?.id ?? null
    const ids = this.c.objects.selected
    if (ids.length !== 1) return null
    return this.c.sheet.charts.some((ch) => ch.id === ids[0]) ? ids[0] : null
  }

  set selected(id: number | null) {
    this.c.objects.selected = id === null ? [] : [id]
  }
  /** The chart last copied, kept whole (its ranges too) for a paste back into Cells. */
  private clip: { token: string; chart: SheetChart } | null = null
  /**
   * A PNG of a chart, when the view has one ready (it draws one as a chart is selected; the copy event can't wait
   * for a picture to be drawn). Null: the copy carries an SVG picture instead.
   */
  picture: ((chart: SheetChart, spec: ChartSpec, size: { w: number; h: number }) => string | null) | null = null

  constructor(private readonly c: SheetController) {}

  /** The selected chart on the sheet shown. */
  active(): SheetChart | null {
    if (this.selected === null) return null
    return this.c.sheet.charts.find((ch) => ch.id === this.selected) ?? null
  }

  select(id: number | null) {
    this.c.objects.select(id)
  }

  /** The spec a chart is drawn with: its own look, with the categories and series read from its ranges now. */
  /** The chart as drawn: its look, its data read from the sheet, and the workbook theme's accent colours. */
  liveSpec(chart: SheetChart): ChartSpec {
    const look = chart.spec as Partial<ChartSpec>
    // A PivotChart: its points from its PivotTable, the outer row fields as the axis' outer levels.
    const pivot = chart.pivot && this.c.pivots.byLink(chart.pivot)
    if (pivot) {
      const data = pivotChartData(pivot.table, pivot.cache, chart.pivot!.sheet)
      const spec = withSheetData(look, data)
      spec.levels = data.levels.length ? data.levels : undefined
      spec.pivotButtons = pivotButtonsShown(pivot.table, pivot.cache, chart.pivot!, spec.legend !== "none")
      return { ...spec, accents: themeAccents(this.c.wb.theme) }
    }
    const spec = chart.refs
      ? withSheetData(look, readChartData(this.c.wb.sheets, this.c.wb.styles, chart.refs))
      : normalizeChartSpec(look)
    return { ...spec, accents: themeAccents(this.c.wb.theme) }
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

  // ── Move Chart and chart sheets ──

  /**
   * Chart Design › Move Chart › New sheet: the selected chart on a chart sheet of its own named `name`, before the
   * sheet it was on (a chart sheet is renamed). False with Excel's reason in an alert when the name is taken.
   */
  moveToNewSheet(name: string): boolean {
    const chart = this.active()
    const clean = name.trim()
    if (!chart || !clean) return false
    const from = this.c.sheetIndex
    const sheet = this.c.sheet
    if (sheet.chartSheet) {
      const problem = this.c.sheets.rename(from, clean)
      if (problem) this.c.shell.alert(problem)
      return !problem
    }
    const problem = sheetNameProblem(
      clean,
      this.c.wb.sheets.map((s) => s.name),
    )
    if (problem) {
      this.c.shell.alert(problem)
      return false
    }
    const ops: Op[] = [
      { kind: "objects", sheet: from, objects: sheet.objects.filter((o) => o !== chart) },
      { kind: "addSheet", sheet: from, name: clean, snapshot: chartSheetSnapshot(clean, chart) },
    ]
    this.moveOps(ops, from)
    return true
  }

  /**
   * Move Chart › Object in: the selected chart onto worksheet `target` (an index), at the top left of what it shows,
   * the size Excel gives a new chart. A chart sheet it leaves goes.
   */
  moveToSheet(target: number): boolean {
    const chart = this.active()
    const to = this.c.wb.sheets[target]
    const from = this.c.sheetIndex
    if (!chart || !to || to.chartSheet || target === from) return false
    const placed: SheetChart = {
      ...copyChart(chart),
      anchor: anchorOf(to, { x: (2 * to.cols.size(0)) / 3, y: to.rows.size(0), w: CHART_WIDTH, h: CHART_HEIGHT }),
    }
    const ops: Op[] = [{ kind: "objects", sheet: target, objects: [...to.objects, placed] }]
    if (this.c.sheet.chartSheet) ops.push({ kind: "removeSheet", sheet: from })
    else ops.push({ kind: "objects", sheet: from, objects: this.c.sheet.objects.filter((o) => o !== chart) })
    const after = this.c.sheet.chartSheet && target > from ? target - 1 : target
    this.moveOps(ops, after)
    this.select(chart.id)
    return true
  }

  /** F11: a chart of the selection (or the data around the active cell) on a new chart sheet before this one. */
  insertSheet(variant = "column"): boolean {
    if (this.c.sheet.chartSheet) return false
    if (this.c.edit) this.c.commitEdit()
    const sheet = this.c.sheet
    const sel = this.c.selection.ranges[0]
    const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
    const range = single ? currentRegion(sheet, this.c.selection.active) : { ...sel }
    const empty = single && sheet.cells.get(range.r0, range.c0)?.value === undefined && range.r0 === range.r1
    const { spec, refs } = this.draft(variant, empty ? null : range)
    const name = nextChartSheetName(this.c.wb.sheets.map((s) => s.name))
    const chart: SheetChart = {
      id: nextChartId(this.c.wb.sheets),
      name: "Chart 1",
      anchor: { from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 0, row: 0, x: 0, y: 0 } },
      spec: spec as unknown as Record<string, unknown>,
      refs,
    }
    const at = this.c.sheetIndex
    this.moveOps([{ kind: "addSheet", sheet: at, name, snapshot: chartSheetSnapshot(name, chart) }], at, "Insert Chart")
    this.c.shell.showTab?.("TabChartDesign")
    return true
  }

  private moveOps(ops: Op[], sheetAfter: number, label = "Move Chart") {
    if (this.c.edit) this.c.commitEdit()
    this.c.transact(ops, { label, before: this.c.selection, sheetAfter })
    this.c.activate(sheetAfter)
  }

  // ── PivotChart Analyze ──

  /** The selected chart when it is a PivotChart whose PivotTable is there. */
  activePivot(): SheetChart | null {
    const chart = this.active()
    return chart?.pivot && this.c.pivots.byLink(chart.pivot) ? chart : null
  }

  /** The selected PivotChart's field buttons (all on unless turned off). */
  pivotButtons(): NonNullable<PivotChartLink["buttons"]> | null {
    const chart = this.activePivot()
    if (!chart) return null
    return chart.pivot!.buttons ?? { filter: true, categories: true, data: true, series: true, visible: true }
  }

  /** PivotChart Analyze › Field Buttons: some of them shown or hidden (`visible` false hides them all). */
  setPivotButtons(change: Partial<NonNullable<PivotChartLink["buttons"]>>): boolean {
    const chart = this.activePivot()
    const now = this.pivotButtons()
    if (!chart || !now) return false
    let next = { ...now, ...change }
    // Showing one kind brings the buttons back when Hide All had hidden them.
    if (!("visible" in change) && Object.values(change).some(Boolean)) next = { ...next, visible: true }
    return this.replace(chart, { ...chart, pivot: { ...chart.pivot!, buttons: next } }, "Field Buttons")
  }

  /** PivotChart Analyze › Refresh: its PivotTable's data read again. */
  refreshPivot(): boolean {
    const chart = this.activePivot()
    const at = chart && this.c.pivots.byLink(chart.pivot!)
    return at ? this.c.pivots.refreshCache(at.table.cacheId) : false
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
    const { spec, refs } = this.draft(variant, empty ? null : range)
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

  /**
   * Insert › PivotChart (and PivotChart Analyze › PivotChart): a chart of PivotTable `link` on this sheet, in the
   * middle of the window. Its points follow the PivotTable; a one-series chart is titled with the series' name.
   */
  insertPivot(link: PivotChartLink, variant = "column"): SheetChart | null {
    if (this.c.edit) this.c.commitEdit()
    const pivot = this.c.pivots.byLink(link)
    if (!pivot) return null
    const sheet = this.c.sheet
    const data = pivotChartData(pivot.table, pivot.cache, link.sheet)
    const spec: ChartSpec = { ...withVariant(defaultChartSpec("column"), variant), categories: [], series: [] }
    spec.title = { show: true, text: data.series.length === 1 ? data.series[0].name : "Chart Title" }
    if (data.series.length === 1 && spec.type !== "pie" && spec.type !== "doughnut") spec.legend = "none"
    const zoom = this.c.zoom || 1
    const x = Math.max(0, this.c.scrollX + (this.c.viewWidth / zoom - CHART_WIDTH) / 2)
    const y = Math.max(0, this.c.scrollY + (this.c.viewHeight / zoom - CHART_HEIGHT) / 2)
    const chart: SheetChart = {
      id: nextChartId(this.c.wb.sheets),
      name: nextChartName(sheet),
      anchor: anchorOf(sheet, { x, y, w: CHART_WIDTH, h: CHART_HEIGHT }),
      spec: spec as unknown as Record<string, unknown>,
      refs: null,
      pivot: { sheet: link.sheet, name: link.name },
    }
    if (!this.commit([...sheet.charts, chart], "Insert PivotChart")) return null
    this.select(chart.id)
    this.c.shell.showTab?.("TabChartDesign")
    return chart
  }

  /**
   * A new chart's look and ranges for data on this sheet (null: an empty chart). Quick Analysis draws its previews
   * from the same, with the data read in (`liveSpec`).
   */
  draft(variant: string, range: Range | null): { spec: ChartSpec; refs: ChartRefs | null } {
    const sheet = this.c.sheet
    const refs = range ? chartRefsFor(sheet, range) : null
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
    return { spec, refs }
  }

  // ── Clipboard ──

  /**
   * Copy (or Cut) of the selected chart: Inkspire's chart markup with a picture inside, and the title as text. Cut
   * removes the chart, as in Excel.
   */
  copy(cut: boolean): { text: string; html: string } | null {
    const chart = this.active()
    if (!chart) return null
    const spec = this.liveSpec(chart)
    const box = chartBox(this.c.sheet, chart.anchor)
    const size = { w: box.w, h: box.h }
    const picture = this.picture?.(chart, spec, size) ?? svgPicture(renderChart(spec, size.w, size.h), size)
    const info: CopiedChartInfo = {
      app: "cells",
      version: 1,
      token: `${Date.now().toString(36)}-${chart.id}`,
      file: this.c.shell.documentPath?.() ?? null,
      sheet: this.c.sheet.name,
      name: chart.name,
      range: this.dataRange(chart),
    }
    this.clip = { token: info.token, chart: copyChart(chart) }
    const html = chartHtml(spec, variantOf(spec), size, picture, info)
    if (cut && !this.c.sheet.chartSheet) this.delete()
    return { text: spec.title?.show && spec.title.text ? spec.title.text : chart.name, html }
  }

  /**
   * Paste of a chart Cells copied: a copy with its own ranges at the active cell (a new name and id), selected.
   * False when the clipboard holds no chart of ours.
   */
  paste(html: string | null | undefined): boolean {
    const info = html ? copiedChartInfo(html) : null
    if (!info || !this.clip || info.token !== this.clip.token) return false
    const sheet = this.c.sheet
    const source = this.clip.chart
    const box = chartBox(sheet, source.anchor)
    const at = this.c.selection.active
    const chart: SheetChart = {
      ...copyChart(source),
      id: nextChartId(this.c.wb.sheets),
      name: nextChartName(sheet),
      anchor: anchorOf(sheet, { x: sheet.cols.offset(at.col), y: sheet.rows.offset(at.row), w: box.w, h: box.h }),
    }
    if (!this.commit([...sheet.charts, chart], "Paste")) return false
    this.select(chart.id)
    return true
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

  /** Alt Text: what a screen reader says for a chart, or that it's only decoration (one undo step). */
  setAltText(id: number, text: string, decorative: boolean) {
    const chart = this.c.sheet.charts.find((x) => x.id === id)
    if (!chart) return
    const next = copyChart(chart)
    const t = text.trim()
    if (t) next.altText = t
    else delete next.altText
    if (decorative) next.decorative = true
    else delete next.decorative
    if (next.altText === chart.altText && next.decorative === chart.decorative) return
    this.replace(chart, next, "Alt Text")
  }

  delete() {
    const chart = this.active()
    // A chart sheet's chart goes with its sheet (Delete Sheet), as in Excel.
    if (!chart || this.c.sheet.chartSheet) return
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

/**
 * The field buttons a PivotChart shows (PivotChart Analyze › Field Buttons; all of them unless turned off). The
 * legend fields' buttons go with the legend, as in Excel.
 */
export function pivotButtonsShown(table: PivotTable, cache: PivotCache, link: PivotChartLink, legend: boolean) {
  const on = link.buttons ?? { filter: true, categories: true, data: true, series: true, visible: true }
  const all = pivotChartButtons(table, cache)
  const shown = <T>(yes: boolean, list: T[]) => (on.visible && yes ? list : [])
  return {
    filters: shown(on.filter, all.filters),
    values: shown(on.data, all.values),
    axis: shown(on.categories, all.axis),
    legend: shown(on.series && legend, all.legend),
  }
}
