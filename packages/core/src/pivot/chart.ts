// PivotCharts: a chart that shows a PivotTable. Its points are the PivotTable's innermost lines, as Excel 365 draws
// them (io's fixtures/excel/pivot-chart.xlsx): the categories are the row lines that have no line inside them (a
// collapsed item is one, its inner levels empty), outer row fields giving the outer levels of the axis; each series is
// an innermost column line (the item, or the value field's name when the Values field is there), or "Total" when the
// PivotTable has neither. Subtotals and grand totals stay out. The refs Excel writes span the PivotTable's label and
// value columns down to its last row; the values come from the PivotTable itself.
import { displayText } from "../format"
import type { ChartRefs } from "../chart"
import { chartRefText } from "../chart"
import { pivotFieldName, type PivotAxisLine } from "./engine"
import { pivotLook } from "./getpivotdata"
import { VALUES_FIELD, type PivotCache, type PivotTable, type PivotValue } from "./model"

export type PivotChartData = {
  /** The innermost label of each category ("" for a collapsed item's inner level). */
  categories: string[]
  /** The outer levels' labels, outermost first, one per category. */
  levels: string[][]
  series: { name: string; values: number[]; texts: string[]; blanks: number[] }[]
  /** Where Excel's c:f point: the label columns, each series' value column and its header cells. */
  refs: ChartRefs
}

/** The lines of an axis that have no line inside them (no subtotals or grand totals). */
function leaves(lines: readonly PivotAxisLine[]): number[] {
  const out: number[] = []
  lines.forEach((l, i) => {
    if (l.t !== "data") return
    const next = lines[i + 1]
    const inside =
      next !== undefined &&
      next.t === "data" &&
      next.items.length > l.items.length &&
      l.items.every((x, j) => next.items[j] === x)
    if (!inside) out.push(i)
  })
  return out
}

const text = (v: PivotValue, fmt?: string): string =>
  v === null || v === "" ? "(blank)" : typeof v === "number" ? displayText(v, fmt ?? "General") : displayText(v, "General")

/** What PivotTable `table` shows as a chart, and the refs to write for it on sheet `sheetName`. */
export function pivotChartData(table: PivotTable, cache: PivotCache, sheetName: string): PivotChartData {
  const { layout, axes, cells } = pivotLook(table, cache)
  const fmtOf = (f: number) => table.fields[f]?.numFmt ?? (cache.fields[f]?.dates ? cache.fields[f].numFmt : undefined)
  const label = (fields: number[], line: PivotAxisLine, level: number) => {
    const f = fields[level]
    const pos = line.items[level]
    return pos === undefined ? "" : text(axes.itemValue(f, pos), fmtOf(f))
  }
  const valueName = (i: number) => table.values[i]?.name ?? ""
  const dataOnRows = table.rows.includes(VALUES_FIELD) && table.values.length > 1
  const dataOnCols = table.cols.includes(VALUES_FIELD) && table.values.length > 1

  // Categories: the innermost row lines (with the value field's name as the innermost level when Values is there).
  let rowLeaves = leaves(axes.rowLines)
  if (!rowLeaves.length) rowLeaves = axes.rowLines.flatMap((l, i) => (l.t === "grand" ? [i] : []))
  const depth = axes.rowFields.length + (dataOnRows ? 1 : 0)
  const paths = rowLeaves.map((i) => {
    const line = axes.rowLines[i]
    if (line.t === "grand") return ["Total"]
    const path = axes.rowFields.map((_, level) => label(axes.rowFields, line, level))
    const own = line.items.length
    const filled = path.slice(0, own)
    if (dataOnRows) filled.push(valueName(line.data))
    while (filled.length < Math.max(depth, 1)) filled.push("")
    return filled
  })
  const width = Math.max(1, ...paths.map((p) => p.length))
  const categories = paths.map((p) => p[width - 1] ?? "")
  const levels = Array.from({ length: width - 1 }, (_, k) => paths.map((p) => p[k] ?? ""))

  // Series: the innermost column lines, or the grand total when there are none.
  let colLeaves = leaves(axes.colLines)
  if (!colLeaves.length) colLeaves = axes.colLines.flatMap((l, i) => (l.t === "grand" ? [i] : []))
  const series = colLeaves.map((ci) => {
    const line = axes.colLines[ci]
    const parts = axes.colFields.slice(0, line.items.length).map((_, level) => label(axes.colFields, line, level))
    if (dataOnCols) parts.push(valueName(line.data))
    const name = parts.length ? parts.join(" - ") : table.values.length === 1 ? "Total" : valueName(line.data)
    const values: number[] = []
    const texts: string[] = []
    const blanks: number[] = []
    rowLeaves.forEach((ri, k) => {
      const at = cells.get(`${ri},${ci}`)
      const v = at ? layout.cells.find((c) => c.row === at.row && c.col === at.col)?.value : undefined
      values.push(typeof v === "number" ? v : 0)
      texts.push(typeof v === "string" ? v : "")
      if (v === undefined || v === null || v === "") blanks.push(k)
    })
    return { name, values, texts, blanks }
  })

  // Excel's refs: from the first line to the table's last row (grand total and all), the label columns and each
  // value column.
  const colOf = new Map<number, number>()
  for (const [key, at] of cells) {
    const ci = Number(key.split(",")[1])
    if (!colOf.has(ci)) colOf.set(ci, at.col)
  }
  const { range } = layout
  const r0 = range.r0 + layout.firstDataRow
  const r1 = Math.max(range.r1, r0)
  const labelCols = Math.max(1, ...layout.rowLabelCols.map((c) => c + 1))
  const headerEnd = range.r0 + Math.max(layout.firstDataRow - 1, 0)
  const refs: ChartRefs = {
    categories: chartRefText(sheetName, { r0, r1, c0: range.c0, c1: range.c0 + labelCols - 1 }),
    series: colLeaves.map((ci) => {
      const col = colOf.get(ci) ?? range.c0 + layout.firstDataCol
      return {
        name: chartRefText(sheetName, { r0: range.r0, r1: headerEnd, c0: col, c1: col }),
        values: chartRefText(sheetName, { r0, r1, c0: col, c1: col }),
      }
    }),
  }
  return { categories, levels, series, refs }
}

/** The PivotTable a PivotChart shows, named as a chart's c:pivotSource does ("[Book1.xlsx]Sheet1!PivotTable1"). */
export function parsePivotSourceName(text: string): { sheet: string; name: string } | null {
  const m = /^(?:\[[^\]]*\])?(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(text.trim())
  if (!m) return null
  return { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), name: m[3] }
}

/** The field names a PivotChart's buttons show: report filters, values, axis (categories) and legend (series). */
export function pivotChartButtons(table: PivotTable, cache: PivotCache) {
  const name = (f: number) => pivotFieldName(table, cache, f)
  return {
    filters: table.pages.map((p) => ({ field: p.field, name: name(p.field) })),
    values: table.values.map((v, index) => ({ index, name: v.name })),
    axis: table.rows.filter((f) => f !== VALUES_FIELD).map((f) => ({ field: f, name: name(f) })),
    legend: table.cols.filter((f) => f !== VALUES_FIELD).map((f) => ({ field: f, name: name(f) })),
  }
}
