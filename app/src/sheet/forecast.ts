import {
  chartRefText,
  contentForInput,
  detectSeasonality,
  displayText,
  etsConfidence,
  etsForecast,
  etsPosition,
  etsSeries,
  etsTime,
  fitEts,
  nextChartId,
  nextTableId,
  nextTableName,
  Sheet,
  uniqueColumnNames,
  type CellWrite,
  type EtsSeries,
  type Op,
  type Range,
  type SheetChart,
  type Table,
} from "@cells/core"
import { defaultChartSpec, withVariant, type ChartSpec } from "@cells/chart"
import { anchorOf, CHART_HEIGHT, CHART_WIDTH } from "./charts"
import type { SheetController } from "./controller"

/**
 * Data › Forecast › Forecast Sheet: a new sheet with the history, a table of forecasts from FORECAST.ETS (with the
 * confidence bounds from FORECAST.ETS.CONFINT) up to the end date, a chart of both, and optionally the model's
 * statistics. Formulas, so the forecast follows edits to the history on the new sheet.
 */

export type ForecastOptions = {
  /** The history: the timeline and values as single columns of the same rows, on one sheet. */
  sheet: number
  timeline: Range
  values: Range
  /** The forecast's last timeline value. */
  end: number
  /** Where the forecast starts (the history's last point unless an earlier one is chosen). */
  start: number
  /** The confidence interval (0.95); null for none. */
  confidence: number | null
  /** 1 detects the season length; 0 none; else the length. */
  seasonality: number
  stats: boolean
  /** 1 fills missing points by interpolation, 0 with zeros. */
  completion: number
  /** How points on the same date combine: 1 Average … 7 Sum. */
  aggregation: number
  chart: "line" | "column"
}

export type ForecastHistory = {
  /** The columns' headings (Timeline and Values when the ranges have none). */
  names: [string, string]
  header: boolean
  /** Points in timeline order: the timeline value and the value (null when missing). */
  points: { t: number; v: number | null }[]
  timeFormat: string
  timeStyle: number
  valueStyle: number
}

/** The history in the ranges, or a problem to show. */
export function readHistory(controller: SheetController, o: Pick<ForecastOptions, "sheet" | "timeline" | "values">) {
  const sheet = controller.wb.sheets[o.sheet]
  const { timeline: t, values: v } = o
  if (t.c0 !== t.c1 || v.c0 !== v.c1) return { problem: "The timeline and the values must each be one column." }
  if (t.r1 - t.r0 !== v.r1 - v.r0) return { problem: "The timeline range must be the same size as the values range." }
  const header = typeof sheet.cells.value(v.r0, v.c0) === "string" || typeof sheet.cells.value(t.r0, t.c0) === "string"
  const text = (row: number, col: number, fallback: string) => {
    const value = sheet.cells.value(row, col)
    return header && value !== undefined && value !== "" ? String(value) : fallback
  }
  const names = uniqueColumnNames([text(t.r0, t.c0, "Timeline"), text(v.r0, v.c0, "Values")]) as [string, string]
  const points: { t: number; v: number | null }[] = []
  for (let i = header ? 1 : 0; i <= t.r1 - t.r0; i++) {
    const time = sheet.cells.value(t.r0 + i, t.c0)
    const value = sheet.cells.value(v.r0 + i, v.c0)
    if (time === undefined && value === undefined) continue
    if (typeof time !== "number") return { problem: "The timeline range must contain numbers or dates." }
    if (value !== undefined && typeof value !== "number") return { problem: "The values range must contain numbers." }
    points.push({ t: time, v: value ?? null })
  }
  if (points.length < 3) return { problem: "There must be at least three points in the history to forecast." }
  points.sort((a, b) => a.t - b.t)
  const first = header ? t.r0 + 1 : t.r0
  const timeStyle = sheet.styleAt(first, t.c0)
  const history: ForecastHistory = {
    names,
    header,
    points,
    timeFormat: controller.wb.styles.get(timeStyle).numberFormat ?? "General",
    timeStyle,
    valueStyle: sheet.styleAt(first, v.c0),
  }
  return { history }
}

/** The series of the history up to `upTo` (all of it when omitted), or a problem. */
export function historySeries(
  h: ForecastHistory,
  o: Pick<ForecastOptions, "completion" | "aggregation">,
  upTo = Infinity,
): { series: EtsSeries } | { problem: string } {
  const shown = h.points.filter((p) => p.t <= upTo)
  const s = etsSeries(
    shown.map((p) => p.t),
    shown.map((p) => p.v),
    o.completion,
    o.aggregation,
  )
  if (typeof s === "string")
    return { problem: "The timeline needs a constant step between its points (missing points are filled in)." }
  return { series: s }
}

/** The timeline values of the forecast after the history, up to `end` (and at most `limit` of them). */
export function futureTimes(full: EtsSeries, end: number, limit = 10000): number[] {
  const out: number[] = []
  for (let k = 1; k <= limit; k++) {
    const t = etsTime(full, full.values.length - 1 + k)
    if (t > end + 1e-9) break
    out.push(t)
  }
  return out
}

/** The forecast's points for the dialog's preview chart, as the sheet will work them out. */
export function forecastPreview(h: ForecastHistory, o: ForecastOptions): { spec: ChartSpec } | { problem: string } {
  const full = historySeries(h, o)
  if ("problem" in full) return full
  const base = historySeries(h, o, o.start)
  if ("problem" in base) return base
  const s = base.series
  const m = o.seasonality === 1 ? detectSeasonality(s.values) : o.seasonality
  const model = fitEts(s.values, m > 1 && m * 2 <= s.values.length ? m : 0)
  const future = futureTimes(full.series, o.end)
  const times = [...h.points.map((p) => p.t), ...future]
  const startIndex = lastAtOrBefore(h.points, o.start)
  const forecast: (number | null)[] = []
  const lower: (number | null)[] = []
  const upper: (number | null)[] = []
  times.forEach((t, i) => {
    if (i < startIndex) return (forecast.push(null), lower.push(null), upper.push(null))
    if (i === startIndex) {
      const v = h.points[i].v
      return (forecast.push(v), lower.push(v), upper.push(v))
    }
    const step = etsPosition(s, t) - (s.values.length - 1)
    const f = etsForecast(model, step)
    const c = o.confidence ? etsConfidence(model, step, o.confidence) : 0
    forecast.push(f)
    lower.push(f - c)
    upper.push(f + c)
  })
  const spec = chartSpec(o.chart)
  spec.categories = times.map((t) => displayText(t, h.timeFormat))
  const nums = (xs: (number | null)[]) => ({
    values: xs.map((x) => x ?? 0),
    blanks: xs.flatMap((x, i) => (x === null ? [i] : [])),
  })
  spec.series = [
    { name: h.names[1], ...nums([...h.points.map((p) => p.v), ...future.map(() => null)]) },
    { name: `Forecast(${h.names[1]})`, ...nums(forecast) },
    ...(o.confidence
      ? [
          { name: `Lower Confidence Bound(${h.names[1]})`, ...nums(lower) },
          { name: `Upper Confidence Bound(${h.names[1]})`, ...nums(upper) },
        ]
      : []),
  ]
  return { spec }
}

function lastAtOrBefore(points: { t: number }[], t: number): number {
  let at = 0
  points.forEach((p, i) => {
    if (p.t <= t + 1e-9) at = i
  })
  return at
}

function chartSpec(kind: "line" | "column"): ChartSpec {
  const spec = withVariant(defaultChartSpec(kind === "line" ? "line" : "column"), kind === "line" ? "line" : "column")
  spec.title = { show: false, text: "" }
  spec.legend = "bottom"
  return spec
}

/** "Sheet2": the first SheetN name not taken. */
function freeSheetName(controller: SheetController): string {
  const taken = new Set(controller.wb.sheets.map((s) => s.name.toLowerCase()))
  for (let n = 1; ; n++) if (!taken.has(`sheet${n}`)) return `Sheet${n}`
}

/** The operations that make the forecast sheet before the active one, and where it is; or a problem to show. */
export function forecastSheetOps(
  controller: SheetController,
  h: ForecastHistory,
  o: ForecastOptions,
): { ops: Op[]; at: number; table: Range } | { problem: string } {
  const full = historySeries(h, o)
  if ("problem" in full) return full
  const last = h.points[h.points.length - 1].t
  if (o.start < h.points[0].t || o.start > last) return { problem: "The forecast start must be within the timeline." }
  if (o.end <= o.start) return { problem: "The forecast end must be after the forecast start." }
  if (o.seasonality > 1 && o.seasonality * 2 > h.points.length)
    return { problem: "The seasonality is too long for the history: it must fit at least twice." }
  const future = futureTimes(full.series, o.end)
  if (!future.length && o.start >= last) return { problem: "The forecast end must be after the end of the history." }
  const styles = controller.wb.styles
  const at = controller.sheetIndex
  const name = freeSheetName(controller)
  const startRow = 1 + lastAtOrBefore(h.points, o.start)
  const lastRow = h.points.length + future.length
  const cols = o.confidence ? 5 : 3
  const heads = [
    h.names[0],
    h.names[1],
    `Forecast(${h.names[1]})`,
    ...(o.confidence ? [`Lower Confidence Bound(${h.names[1]})`, `Upper Confidence Bound(${h.names[1]})`] : []),
  ]
  const cells: CellWrite[] = []
  const put = (row: number, col: number, text: string | number | null, style = 0) => {
    if (text === null) return
    if (typeof text === "number") {
      cells.push({ row, col, data: { value: text, ...(style ? { style } : {}) } })
      return
    }
    const data = contentForInput(styles, style ? { style } : undefined, text)
    if (data) cells.push({ row, col, data })
  }
  heads.forEach((head, col) => cells.push({ row: 0, col, data: { value: head } }))
  h.points.forEach((p, i) => {
    put(i + 1, 0, p.t, h.timeStyle)
    put(i + 1, 1, p.v, h.valueStyle)
  })
  future.forEach((t, i) => put(h.points.length + 1 + i, 0, t, h.timeStyle))
  const values = `$B$2:$B$${startRow + 1}`
  const timeline = `$A$2:$A$${startRow + 1}`
  const tail = `${o.seasonality},${o.completion}${o.aggregation !== 1 ? `,${o.aggregation}` : ""}`
  const startValue = h.points[startRow - 1].v
  for (let row = startRow; row <= lastRow; row++) {
    const r = row + 1
    if (row === startRow) {
      for (let col = 2; col < cols; col++) put(row, col, startValue, h.valueStyle)
      continue
    }
    put(row, 2, `=FORECAST.ETS(A${r},${values},${timeline},${tail})`, h.valueStyle)
    if (o.confidence) {
      const confint = `FORECAST.ETS.CONFINT(A${r},${values},${timeline},${o.confidence},${tail})`
      put(row, 3, `=C${r}-${confint}`, h.valueStyle)
      put(row, 4, `=C${r}+${confint}`, h.valueStyle)
    }
  }
  const table: Range = { r0: 0, c0: 0, r1: lastRow, c1: cols - 1 }
  const tableNames = controller.wb.names.all().map((n) => n.name)
  const t: Table = {
    id: nextTableId(controller.wb.sheets),
    name: nextTableName(controller.wb.sheets, tableNames),
    range: table,
    header: true,
    totals: false,
    columns: heads.map((n) => ({ name: n })),
    style: "TableStyleMedium2",
    firstColumn: false,
    lastColumn: false,
    rowStripes: true,
    columnStripes: false,
    autoFilter: { range: { ...table }, columns: {} },
  }
  // Statistics beside the table.
  let chartCol = cols + 1
  if (o.stats) {
    const sc = cols + 1
    put(0, sc, "Statistic")
    put(0, sc + 1, "Value")
    ;["Alpha", "Beta", "Gamma", "MASE", "SMAPE", "MAE", "RMSE"].forEach((label, i) => {
      put(i + 1, sc, label)
      put(i + 1, sc + 1, `=FORECAST.ETS.STAT(${values},${timeline},${i + 1},${tail})`)
    })
    chartCol = sc + 3
  }
  // The chart, from the table's columns, to the right.
  const quoted = (range: Range) => chartRefText(name, range)
  const column = (col: number): Range => ({ r0: 1, c0: col, r1: lastRow, c1: col })
  const head = (col: number): Range => ({ r0: 0, c0: col, r1: 0, c1: col })
  const plain = new Sheet(name)
  const x = plain.cols.offset(chartCol)
  const chart: SheetChart = {
    id: nextChartId(controller.wb.sheets),
    name: "Chart 1",
    anchor: anchorOf(plain, { x, y: plain.rows.offset(1), w: CHART_WIDTH, h: CHART_HEIGHT }),
    spec: { ...chartSpec(o.chart), categories: [], series: [] } as unknown as Record<string, unknown>,
    refs: {
      categories: quoted(column(0)),
      series: Array.from({ length: cols - 1 }, (_, i) => ({
        name: quoted(head(i + 1)),
        values: quoted(column(i + 1)),
      })),
    },
  }
  const ops: Op[] = [
    { kind: "addSheet", sheet: at, name },
    { kind: "setCells", sheet: at, cells },
    { kind: "tables", sheet: at, tables: [t] },
    { kind: "charts", sheet: at, charts: [chart] },
  ]
  if (o.stats)
    ops.push({
      kind: "format",
      sheet: at,
      ranges: [{ r0: 0, c0: cols + 1, r1: 0, c1: cols + 2 }],
      patch: { bold: true },
    })
  return { ops, at, table }
}

/** Forecast Sheet › Create: the sheet made and shown, its columns fitted. Null when done, else the problem. */
export function createForecastSheet(
  controller: SheetController,
  h: ForecastHistory,
  o: ForecastOptions,
): string | null {
  if (controller.edit) controller.commitEdit()
  const result = forecastSheetOps(controller, h, o)
  if ("problem" in result) return result.problem
  const { ops, at, table } = result
  const entry = controller.transact(ops, { label: "Forecast Sheet", before: controller.selection, sheetAfter: at })
  if (!entry) return null
  controller.activate(at)
  const fit = controller.columnFit
  const widths = []
  for (let col = table.c0; fit && col <= table.c1 + (o.stats ? 3 : 0); col++) {
    const size = fit(col, { r0: table.r0, r1: table.r1 })
    if (size !== null && size > controller.sheet.cols.size(col)) widths.push({ index: col, size, hidden: false })
  }
  if (widths.length) controller.wb.amend(entry, [{ kind: "axis", sheet: at, axis: "cols", entries: widths }])
  return null
}
