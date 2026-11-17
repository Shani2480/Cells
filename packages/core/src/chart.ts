import { columnName, MAX_COLS, MAX_ROWS, type Range } from "./address"
import { displayText } from "./format"
import { parseRefText } from "./formula/parser"
import { quoteSheet } from "./formula/print"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"

/**
 * Charts on a sheet, as the workbook keeps them: where each sits (anchored to cells, so it moves and sizes with
 * them, Excel's default) and which ranges feed its series. How it looks (type, elements, colours) is the chart spec,
 * which the core keeps as data and the chart package reads.
 */

/** A corner of a chart: a cell and an offset into it, in pixels at 100%. */
export type ChartCorner = { row: number; col: number; x: number; y: number }
export type ChartAnchor = { from: ChartCorner; to: ChartCorner }

/**
 * The ranges a chart reads, as Excel writes them in a chart (c:f): "Sheet1!$B$1", "'My data'!$A$2:$A$5". Null:
 * typed into the chart itself (or no categories).
 */
export type ChartSeriesRefs = { name: string | null; values: string; bubbles?: string | null }
export type ChartRefs = { categories: string | null; series: ChartSeriesRefs[] }

export type SheetChart = {
  id: number
  /** "Chart 1", as Excel names drawing objects. */
  name: string
  anchor: ChartAnchor
  /** The chart spec (see @cells/chart's ChartSpec): the series' values in it are refreshed from `refs`. */
  spec: Record<string, unknown>
  refs: ChartRefs | null
  /** Alt Text: what a screen reader says for the chart. */
  altText?: string
  /** Alt Text › Mark as decorative: a screen reader skips the chart. */
  decorative?: boolean
  /** Inside a group: where it sits in the group (pixels of the group's own size). */
  box?: { x: number; y: number; w: number; h: number }
}

const corner = (c: ChartCorner): ChartCorner => ({ ...c })

export function copyChart(c: SheetChart): SheetChart {
  return {
    ...c,
    anchor: { from: corner(c.anchor.from), to: corner(c.anchor.to) },
    spec: structuredCopy(c.spec),
    refs: c.refs ? { categories: c.refs.categories, series: c.refs.series.map((s) => ({ ...s })) } : null,
  }
}

export const copyCharts = (charts: readonly SheetChart[]) => charts.map(copyChart)

function structuredCopy<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

/** Charts follow inserted and deleted rows and columns: a corner in deleted lines moves to where they were. */
export function shiftCharts(charts: SheetChart[], axis: "rows" | "cols", at: number, count: number) {
  const key = axis === "rows" ? "row" : "col"
  const offset = axis === "rows" ? "y" : "x"
  for (const chart of charts)
    for (const c of [chart.anchor.from, chart.anchor.to]) {
      if (count > 0) {
        if (c[key] >= at) c[key] += count
      } else if (c[key] >= at - count) c[key] += count
      else if (c[key] >= at) {
        c[key] = at
        c[offset] = 0
      }
    }
}

/** Every ref of a chart, for rewriting: `map` gets each formula text and returns the new one (null: unchanged). */
export function rewriteChartRefs(chart: SheetChart, map: (ref: string) => string | null): boolean {
  if (!chart.refs) return false
  let changed = false
  const go = (ref: string | null | undefined) => {
    if (!ref) return ref
    const next = map(ref)
    if (next === null || next === ref) return ref
    changed = true
    return next
  }
  chart.refs.categories = go(chart.refs.categories) ?? null
  for (const s of chart.refs.series) {
    s.name = go(s.name) ?? null
    s.values = go(s.values) ?? s.values
    if (s.bubbles !== undefined) s.bubbles = go(s.bubbles) ?? null
  }
  return changed
}

/** The next free chart id across the workbook (charts, shapes and pictures share ids, as in a drawing). */
export function nextChartId(sheets: readonly { charts: SheetChart[]; objects?: readonly { id: number }[] }[]): number {
  let max = 0
  const visit = (o: { id: number; items?: readonly { id: number }[] }) => {
    max = Math.max(max, o.id)
    o.items?.forEach(visit)
  }
  for (const s of sheets) (s.objects ?? s.charts).forEach(visit)
  return max + 1
}

export function nextChartName(sheet: { charts: SheetChart[] }): string {
  let n = 1
  const taken = new Set(sheet.charts.map((c) => c.name.toLowerCase()))
  while (taken.has(`chart ${n}`)) n++
  return `Chart ${n}`
}

// ── Reading a chart's data from its ranges ──────────────────────────────────────────────────────────────────────

/** A ref's sheet and range (null when the sheet is gone or the text isn't one range). */
export function chartRefRange(sheets: readonly Sheet[], ref: string): { sheet: Sheet; range: Range } | null {
  const node = parseRefText(ref.replace(/^=/, "").trim())
  if (!node || node.invalid || !node.sheet || node.sheetTo) return null
  const name = node.sheet.toLowerCase()
  const sheet = sheets.find((s) => s.name.toLowerCase() === name)
  if (!sheet) return null
  const a = node.from
  const b = node.to ?? node.from
  const r0 = a.row ?? 0
  const r1 = b.row ?? (a.row === null ? MAX_ROWS - 1 : r0)
  const c0 = a.col ?? 0
  const c1 = b.col ?? (a.col === null ? MAX_COLS - 1 : c0)
  return { sheet, range: { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) } }
}

/** A range's cells in order (along its longer side), clipped to the cells in use. */
function refCells(sheets: readonly Sheet[], ref: string | null | undefined) {
  if (!ref) return null
  const at = chartRefRange(sheets, ref)
  if (!at) return null
  const { sheet, range } = at
  const used = sheet.cells.usedBounds()
  const r1 = used ? Math.min(range.r1, Math.max(range.r0, used.r1)) : range.r0
  const c1 = used ? Math.min(range.c1, Math.max(range.c0, used.c1)) : range.c0
  const out: { sheet: Sheet; row: number; col: number }[] = []
  for (let row = range.r0; row <= r1; row++)
    for (let col = range.c0; col <= c1; col++)
      if (!sheet.rows.isHidden(row) && !sheet.cols.isHidden(col)) out.push({ sheet, row, col })
  return out
}

/**
 * The category labels: one per row of the range (per column when it runs across), the texts of a row of several
 * label columns joined, as Excel's multi-level labels read.
 */
function categoryTexts(
  sheets: readonly Sheet[],
  ref: string | null,
  shown: (c: { sheet: Sheet; row: number; col: number }) => string,
): string[] | null {
  const cells = refCells(sheets, ref)
  if (!cells) return null
  const at = chartRefRange(sheets, ref!)!
  const down = at.range.r1 > at.range.r0 || at.range.c0 === at.range.c1
  const groups = new Map<number, string[]>()
  for (const c of cells) {
    const key = down ? c.row : c.col
    const text = shown(c)
    const list = groups.get(key) ?? []
    if (text) list.push(text)
    groups.set(key, list)
  }
  return [...groups.values()].map((parts) => parts.join(" "))
}

export type ChartData = {
  categories: string[] | null
  /** `blanks`: the points whose cells are empty, which a line chart leaves as gaps (Excel's default). */
  series: { name: string | null; values: number[]; texts: string[]; blanks?: number[] }[]
}

/**
 * What a chart's ranges hold now: the categories as the cells show them, each series' name and its numbers (text and
 * blanks count as 0, as Excel plots them; their texts are kept for maps). Hidden rows and columns are left out, as
 * Excel does by default.
 */
export function readChartData(sheets: readonly Sheet[], styles: StyleTable, refs: ChartRefs): ChartData {
  const shown = (c: { sheet: Sheet; row: number; col: number }) => {
    const value = c.sheet.cells.get(c.row, c.col)?.value
    return displayText(value, styles.get(c.sheet.styleAt(c.row, c.col)).numberFormat)
  }
  const categories = categoryTexts(sheets, refs.categories, shown)
  const series = refs.series.map((s) => {
    const cells = refCells(sheets, s.values) ?? []
    const name = refCells(sheets, s.name)
    const out: ChartData["series"][number] = {
      name: name ? name.map(shown).join(" ") : null,
      values: cells.map((c) => {
        const v = c.sheet.cells.get(c.row, c.col)?.value
        return typeof v === "number" ? v : 0
      }),
      texts: cells.map((c) => {
        const v = c.sheet.cells.get(c.row, c.col)?.value
        return typeof v === "string" ? v : ""
      }),
    }
    const blanks = cells.flatMap((c, i) => (c.sheet.cells.get(c.row, c.col)?.value === undefined ? [i] : []))
    return blanks.length ? { ...out, blanks } : out
  })
  return { categories, series }
}

// ── A chart's ranges from the data selected ─────────────────────────────────────────────────────────────────────

/** "Sheet1!$B$2:$B$5", as Excel writes a series' range. */
export function chartRefText(sheetName: string, range: Range): string {
  const cell = (row: number, col: number) => `$${columnName(col)}$${row + 1}`
  const one = range.r0 === range.r1 && range.c0 === range.c1
  return `${quoteSheet(sheetName)}!${cell(range.r0, range.c0)}${one ? "" : ":" + cell(range.r1, range.c1)}`
}

/**
 * The ranges Excel's Insert Chart takes from data: series down the columns when there are more rows than columns
 * (across the rows otherwise, or as `byRow` says); a first row with text, or an empty corner, gives the series' names,
 * and the leading columns of text give the categories.
 */
export function chartRefsFor(sheet: Sheet, range: Range, byRow?: boolean): ChartRefs {
  const v = (row: number, col: number) => sheet.cells.get(row, col)?.value
  const isLabel = (x: unknown) => x !== undefined && x !== "" && typeof x !== "number"
  const corner = v(range.r0, range.c0)
  const wide = range.c1 > range.c0
  const tall = range.r1 > range.r0
  const empty = corner === undefined || corner === ""
  let topLabels = tall && empty
  for (let c = range.c0 + (wide ? 1 : 0); c <= range.c1 && tall; c++) if (isLabel(v(range.r0, c))) topLabels = true
  // Leading columns of text under the header (Region, Product) are all category labels, not series.
  const first = range.r0 + (topLabels ? 1 : 0)
  const textColumn = (c: number) => {
    let label = false
    for (let r = first; r <= range.r1; r++) {
      const x = v(r, c)
      if (typeof x === "number") return false
      if (isLabel(x)) label = true
    }
    return label
  }
  let labelCols = 0
  while (range.c0 + labelCols < range.c1 && textColumn(range.c0 + labelCols)) labelCols++
  if (!labelCols && wide && empty) labelCols = 1
  const leftLabels = labelCols > 0
  const r0 = first
  const c0 = range.c0 + labelCols
  const rows = range.r1 - r0 + 1
  const cols = range.c1 - c0 + 1
  const across = byRow ?? rows < cols
  const ref = (r: Range) => chartRefText(sheet.name, r)
  if (!across) {
    return {
      categories: leftLabels ? ref({ r0, r1: range.r1, c0: range.c0, c1: c0 - 1 }) : null,
      series: Array.from({ length: Math.max(cols, 0) }, (_, i) => ({
        name: topLabels ? ref({ r0: range.r0, r1: range.r0, c0: c0 + i, c1: c0 + i }) : null,
        values: ref({ r0, r1: range.r1, c0: c0 + i, c1: c0 + i }),
      })),
    }
  }
  return {
    categories: topLabels ? ref({ r0: range.r0, r1: range.r0, c0, c1: range.c1 }) : null,
    series: Array.from({ length: Math.max(rows, 0) }, (_, i) => ({
      name: leftLabels ? ref({ r0: r0 + i, r1: r0 + i, c0: range.c0, c1: c0 - 1 }) : null,
      values: ref({ r0: r0 + i, r1: r0 + i, c0, c1: range.c1 }),
    })),
  }
}

/** The block of cells a chart's ranges cover (null when they aren't all on one sheet). */
export function chartRefsBox(sheets: readonly Sheet[], refs: ChartRefs): { sheet: Sheet; range: Range } | null {
  const parts = [refs.categories, ...refs.series.flatMap((s) => [s.name, s.values])].filter((r): r is string => !!r)
  const at = parts.map((p) => chartRefRange(sheets, p))
  const first = at[0]
  if (!first || at.some((a) => !a || a.sheet !== first.sheet)) return null
  const range = { ...first.range }
  for (const a of at) {
    range.r0 = Math.min(range.r0, a!.range.r0)
    range.r1 = Math.max(range.r1, a!.range.r1)
    range.c0 = Math.min(range.c0, a!.range.c0)
    range.c1 = Math.max(range.c1, a!.range.c1)
  }
  return { sheet: first.sheet, range }
}

/** Whether a chart's series run across rows (each series a row of cells). */
export function chartSeriesInRows(sheets: readonly Sheet[], refs: ChartRefs): boolean {
  const first = chartRefRange(sheets, refs.series[0]?.values ?? "")
  return first ? first.range.r0 === first.range.r1 && first.range.c1 > first.range.c0 : false
}

/** Switch Row/Column: the same block of data with the series taken the other way. */
export function switchChartRefs(sheets: readonly Sheet[], refs: ChartRefs): ChartRefs | null {
  const box = chartRefsBox(sheets, refs)
  return box ? chartRefsFor(box.sheet, box.range, !chartSeriesInRows(sheets, refs)) : null
}

/**
 * A copied sheet's charts (Move or Copy): new ids, and ranges on the sheet copied now read the copy, as in Excel.
 * Ranges on other sheets stay.
 */
export function copiedSheetCharts(
  charts: readonly SheetChart[],
  sheets: readonly Sheet[],
  from: Sheet,
  copyName: string,
): SheetChart[] {
  let id = nextChartId(sheets)
  return charts.map((c) => {
    const chart = copyChart(c)
    chart.id = id++
    rewriteChartRefs(chart, (ref) => {
      const at = chartRefRange(sheets, ref)
      return at?.sheet === from ? chartRefText(copyName, at.range) : null
    })
    return chart
  })
}
