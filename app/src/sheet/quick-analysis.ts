import {
  chartRefsFor,
  columnName,
  currentRegion,
  DEFAULT_TABLE_STYLE,
  guessHeader,
  isDateFormat,
  rangeName,
  type CellWrite,
  type CFRule,
  type Op,
  type Range,
  type SparklineType,
  type UndoEntry,
} from "@cells/core"
import type { ChartSpec } from "@cells/chart"
import { CF_PRESETS, COLOR_SCALES, DATA_BAR_COLORS } from "./conditional"
import { anchorOf, CHART_HEIGHT, CHART_WIDTH } from "./charts"
import type { SheetController } from "./controller"

/**
 * Quick Analysis (Ctrl+Q, the button at the corner of the selection, the cell menu): Excel's five tabs of one-click
 * analyses for the data selected. Hovering a choice previews it on the sheet (Excel's live preview); a click makes
 * it as one undo step.
 *
 * Built from Excel 365's Quick Analysis as remembered (no reference screenshot in the repo): the tabs, their
 * choices for numbers, text and dates, and the totals' formulas. PivotTables wait for Part 10.
 */

export type QuickTab = "formatting" | "charts" | "totals" | "tables" | "sparklines"

export const QUICK_TABS: { id: QuickTab; label: string; text: string }[] = [
  {
    id: "formatting",
    label: "Formatting",
    text: "Conditional Formatting uses rules to highlight interesting data.",
  },
  { id: "charts", label: "Charts", text: "Recommended Charts help you visualize data." },
  { id: "totals", label: "Totals", text: "Formulas automatically calculate totals for you." },
  { id: "tables", label: "Tables", text: "Tables help you sort, filter, and summarize data." },
  { id: "sparklines", label: "Sparklines", text: "Sparklines are mini charts placed in single cells." },
]

/** One choice on a tab. `id`: "cf:dataBar", "chart:column", "total:below:sum", "table", "spark:line"… */
export type QuickItem = { id: string; label: string; tip: string; disabled?: string }

export type TotalKind = "sum" | "average" | "count" | "percent" | "running"
export type TotalSide = "below" | "right"

const TOTAL_LABEL: Record<TotalKind, string> = {
  sum: "Sum",
  average: "Average",
  count: "Count",
  percent: "% Total",
  running: "Running Total",
}

/** Excel's message when a total or sparkline would go over cells that hold something. */
export const REPLACE_QUESTION = "There's already data here. Do you want to replace it?"

/** What kind of data the range holds, which picks the Formatting tab's choices (as Excel's). */
export type DataKind = "numbers" | "text" | "dates"

/** The selected block read once: its header row, the rows of data, and which columns hold numbers. */
export type QuickShape = {
  range: Range
  header: boolean
  /** The data rows (under the header). */
  top: number
  /** Columns with a number in a data row, left to right. */
  numberCols: number[]
  kind: DataKind
  /** All numbers in the data are zero or more (a pie chart makes sense). */
  positive: boolean
}

/** Workbooks showing a preview: their changes don't count as edits (unsaved changes, AutoRecover). */
export const previewing = new WeakSet<object>()

export class QuickAnalysis {
  /** The preview on the sheet: the step that takes it back. */
  private shown: UndoEntry<unknown> | null = null

  /** Opens the gallery under the selection (set by the view that draws it). */
  openGallery: (() => void) | null = null

  constructor(private readonly c: SheetController) {}

  /** Ctrl+Q and the cell menu: the gallery for the selection (or the data around the active cell). */
  show(): boolean {
    if (!this.prepare()) return false
    this.openGallery?.()
    return true
  }

  /** The block Quick Analysis works on: the selection, cut to the used cells; null when it holds nothing. */
  range(): Range | null {
    const sel = this.c.selection.ranges[0]
    const used = this.c.sheet.cells.usedBounds()
    if (!sel || !used) return null
    const r: Range = {
      r0: Math.max(sel.r0, used.r0),
      c0: Math.max(sel.c0, used.c0),
      r1: Math.min(sel.r1, used.r1),
      c1: Math.min(sel.c1, used.c1),
    }
    if (r.r0 > r.r1 || r.c0 > r.c1) return null
    let any = false
    this.c.sheet.cells.forEach(r, (_row, _col, d) => {
      if (d.value !== undefined && d.value !== "") any = true
    })
    return any ? r : null
  }

  /**
   * Whether the button shows at the selection's corner: Excel shows it for one block of two or more cells with
   * data, while nothing is being typed.
   */
  available(): boolean {
    const ranges = this.c.selection.ranges
    if (ranges.length !== 1 || this.c.edit) return false
    const s = ranges[0]
    if (s.r0 === s.r1 && s.c0 === s.c1) return false
    // Asked on every redraw (scrolling too): read the cells once per selection and change.
    const key = this.availableFor
    if (key && key.selection === this.c.selection && key.version === this.c.wb.version && key.sheet === this.c.sheet)
      return key.answer
    const answer = this.range() !== null
    this.availableFor = { selection: this.c.selection, version: this.c.wb.version, sheet: this.c.sheet, answer }
    return answer
  }

  private availableFor: { selection: object; version: number; sheet: object; answer: boolean } | null = null

  /**
   * Ctrl+Q: on one cell, Excel selects the data around it first. Returns whether there's data to analyse (the
   * shell then opens the gallery).
   */
  prepare(): boolean {
    if (this.c.edit) return false
    const ranges = this.c.selection.ranges
    const s = ranges[0]
    if (ranges.length === 1 && s.r0 === s.r1 && s.c0 === s.c1) {
      const region = currentRegion(this.c.sheet, this.c.selection.active)
      if (region.r0 === region.r1 && region.c0 === region.c1) return false
      const active = this.c.selection.active
      this.c.setSelection(
        {
          ranges: [{ ...region }],
          active,
          activeRange: 0,
          anchor: { row: region.r0, col: region.c0 },
          focus: { row: region.r1, col: region.c1 },
        },
        null,
      )
    }
    return this.range() !== null
  }

  shape(): QuickShape | null {
    const range = this.range()
    if (!range) return null
    const sheet = this.c.sheet
    const styles = this.c.wb.styles
    const header = guessHeader(sheet, styles, range)
    const top = header ? range.r0 + 1 : range.r0
    const counts = { numbers: 0, text: 0, dates: 0 }
    const cols = new Set<number>()
    let positive = true
    sheet.cells.forEach({ ...range, r0: top }, (row, col, d) => {
      const v = d.value
      if (typeof v === "number") {
        cols.add(col)
        if (v < 0) positive = false
        const code = styles.get(sheet.styleAt(row, col)).numberFormat
        counts[code && isDateFormat(code) ? "dates" : "numbers"]++
      } else if (typeof v === "string" && v !== "") counts.text++
    })
    const kind: DataKind =
      counts.dates > counts.numbers && counts.dates >= counts.text
        ? "dates"
        : counts.numbers + counts.dates >= counts.text
          ? "numbers"
          : "text"
    return { range, header, top, numberCols: [...cols].sort((a, b) => a - b), kind, positive }
  }

  // ── The tabs' choices ──

  formatting(shape: QuickShape): QuickItem[] {
    const clear: QuickItem = { id: "cf:clear", label: "Clear Format", tip: "Clear rules from the selected cells." }
    if (shape.kind === "text")
      return [
        { id: "cf:text", label: "Text Contains", tip: "Highlight cells that contain specific text." },
        { id: "cf:duplicate", label: "Duplicate Values", tip: "Highlight values that appear more than once." },
        { id: "cf:unique", label: "Unique Values", tip: "Highlight values that appear only once." },
        { id: "cf:equal", label: "Equal To", tip: "Highlight cells that are equal to a value." },
        clear,
      ]
    if (shape.kind === "dates")
      return [
        { id: "cf:lastMonth", label: "Last Month", tip: "Highlight dates in the last month." },
        { id: "cf:lastWeek", label: "Last Week", tip: "Highlight dates in the last week." },
        { id: "cf:greaterThan", label: "Greater Than", tip: "Highlight dates after a date." },
        { id: "cf:lessThan", label: "Less Than", tip: "Highlight dates before a date." },
        { id: "cf:equal", label: "Equal To", tip: "Highlight cells that are equal to a date." },
        clear,
      ]
    return [
      { id: "cf:dataBar", label: "Data Bars", tip: "A bar in each cell shows its value against the others." },
      { id: "cf:colorScale", label: "Color Scale", tip: "Colors show the high and low values." },
      { id: "cf:iconSet", label: "Icon Set", tip: "Icons show the high, middle and low values." },
      { id: "cf:greaterThan", label: "Greater Than", tip: "Highlight cells greater than a value." },
      { id: "cf:top10", label: "Top 10%", tip: "Highlight the top 10% of the values." },
      clear,
    ]
  }

  /** The rule a Formatting choice makes (with the value Excel suggests); null for the ones that open a dialog. */
  rule(id: string): CFRule | null {
    const style = CF_PRESETS[0].style
    switch (id) {
      case "cf:dataBar":
        return {
          kind: "dataBar",
          min: { type: "automatic" },
          max: { type: "automatic" },
          color: DATA_BAR_COLORS[0].color,
          gradient: true,
        }
      case "cf:colorScale": {
        const [lo, mid, hi] = COLOR_SCALES[0].colors
        return {
          kind: "colorScale",
          points: [
            { at: { type: "min" }, color: lo },
            { at: { type: "percentile", value: "50" }, color: mid },
            { at: { type: "max" }, color: hi },
          ],
        }
      }
      case "cf:iconSet":
        return {
          kind: "iconSet",
          set: "3Arrows",
          reverse: false,
          thresholds: [
            { type: "percent", value: "0" },
            { type: "percent", value: "33" },
            { type: "percent", value: "67" },
          ],
          showValue: true,
        }
      case "cf:greaterThan":
        return { kind: "cellIs", operator: "greaterThan", values: this.c.conditional.suggestion("greaterThan"), style }
      case "cf:lessThan":
        return { kind: "cellIs", operator: "lessThan", values: this.c.conditional.suggestion("lessThan"), style }
      case "cf:top10":
        return { kind: "top", bottom: false, percent: true, rank: 10, style }
      case "cf:duplicate":
        return { kind: "duplicate", unique: false, style }
      case "cf:unique":
        return { kind: "duplicate", unique: true, style }
      case "cf:lastMonth":
        return { kind: "date", period: "lastMonth", style }
      case "cf:lastWeek":
        return { kind: "date", period: "lastWeek", style }
    }
    return null
  }

  /** Charts Excel would recommend for the data, best first (at most five). */
  charts(shape: QuickShape): string[] {
    if (!shape.numberCols.length) return []
    const refs = chartRefsFor(this.c.sheet, shape.range)
    const n = refs.series.length
    let ids: string[]
    if (!refs.categories && n >= 2) ids = ["scatter", "line", "column", "column-stacked", "area"]
    else if (shape.kind === "dates") ids = ["line", "column", "area", "hbar", "column-stacked"]
    else if (n === 1) ids = ["column", "hbar", "pie", "line", "area"]
    else ids = ["column", "column-stacked", "hbar", "line", "area-stacked"]
    if (!shape.positive || n > 1) ids = ids.filter((id) => id !== "pie")
    return ids.slice(0, 5)
  }

  /** A recommended chart drawn from the data, as the preview over the gallery shows it. */
  chartSpec(variant: string): ChartSpec | null {
    const range = this.range()
    if (!range) return null
    const { spec, refs } = this.c.charts.draft(variant, range)
    return this.c.charts.liveSpec({
      id: -1,
      name: "",
      anchor: anchorOf(this.c.sheet, { x: 0, y: 0, w: CHART_WIDTH, h: CHART_HEIGHT }),
      spec: spec as unknown as Record<string, unknown>,
      refs,
    })
  }

  totals(shape: QuickShape): QuickItem[] {
    const kinds: TotalKind[] = shape.numberCols.length ? ["sum", "average", "count", "percent", "running"] : ["count"]
    const items: QuickItem[] = []
    for (const side of ["below", "right"] as TotalSide[])
      for (const kind of kinds)
        items.push({
          id: `total:${side}:${kind}`,
          label: TOTAL_LABEL[kind],
          tip:
            side === "below"
              ? `${TOTAL_LABEL[kind]} of each column, in the row below.`
              : `${TOTAL_LABEL[kind]} of each row, in the column to the right.`,
        })
    return items
  }

  /** The cells a total writes (formulas, labels) and their formats. */
  totalOps(shape: QuickShape, side: TotalSide, kind: TotalKind): { ops: Op[]; cells: CellWrite[] } {
    const { range, top } = shape
    const sheet = this.c.sheet
    const index = this.c.sheetIndex
    const cols = shape.numberCols.length
      ? shape.numberCols
      : Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i)
    const first = cols[0]
    const last = cols[cols.length - 1]
    const fn =
      kind === "average" ? "AVERAGE" : kind === "count" ? (shape.numberCols.length ? "COUNT" : "COUNTA") : "SUM"
    const a1 = (row: number, col: number, absRow = false, absCol = false) =>
      `${absCol ? "$" : ""}${columnName(col)}${absRow ? "$" : ""}${row + 1}`
    const whole = `${a1(top, first, true, true)}:${a1(range.r1, last, true, true)}`
    const cells: CellWrite[] = []
    // A total takes the format of the cell before it (the number above, the header to its left), as Excel's do.
    const styleFrom = (row: number, col: number) => {
      const style = sheet.cells.get(row, col)?.style
      return style ? { style } : {}
    }
    const before = (row: number, col: number) => (side === "below" ? styleFrom(row - 1, col) : styleFrom(row, col - 1))
    const formula = (input: string, row: number, col: number) => ({ ...before(row, col), value: 0, input })
    const label = (text: string, row: number, col: number) => ({ ...before(row, col), value: text })
    if (side === "below") {
      const row = range.r1 + 1
      for (const col of cols) {
        const part = `${a1(top, col)}:${a1(range.r1, col)}`
        const input =
          kind === "percent"
            ? `=SUM(${part})/SUM(${whole})`
            : kind === "running"
              ? `=SUM(${a1(top, first, true, true)}:${a1(range.r1, col)})`
              : `=${fn}(${part})`
        cells.push({ row, col, data: formula(input, row, col) })
      }
      // A column of labels on the left names the row (Excel writes the totals only; the label says what they are).
      if (!cols.includes(range.c0) && range.c0 < first)
        cells.push({ row, col: range.c0, data: label(TOTAL_LABEL[kind], row, range.c0) })
    } else {
      const col = range.c1 + 1
      for (let row = top; row <= range.r1; row++) {
        let number = false
        for (const c of cols) if (typeof sheet.cells.get(row, c)?.value === "number") number = true
        if (!number && shape.numberCols.length) continue
        const part = `${a1(row, first)}:${a1(row, last)}`
        const input =
          kind === "percent"
            ? `=SUM(${part})/SUM(${whole})`
            : kind === "running"
              ? `=SUM(${a1(top, first, true)}:${a1(row, last)})`
              : `=${fn}(${part})`
        cells.push({ row, col, data: formula(input, row, col) })
      }
      if (shape.header) cells.push({ row: range.r0, col, data: label(TOTAL_LABEL[kind], range.r0, col) })
    }
    const spots = cells.map((w) => ({ r0: w.row, c0: w.col, r1: w.row, c1: w.col }))
    const ops: Op[] = [
      { kind: "setCells", sheet: index, cells },
      { kind: "format", sheet: index, ranges: spots, patch: { bold: true } },
    ]
    if (kind === "percent") {
      const numbers = cells.filter((w) => w.data && "input" in w.data)
      ops.push({
        kind: "format",
        sheet: index,
        ranges: numbers.map((w) => ({ r0: w.row, c0: w.col, r1: w.row, c1: w.col })),
        patch: { numberFormat: "0.00%" },
      })
    }
    return { ops, cells }
  }

  tables(shape: QuickShape): QuickItem[] {
    const inTable = this.c.sheet.tables.some(
      (t) =>
        t.range.r0 <= shape.range.r1 &&
        shape.range.r0 <= t.range.r1 &&
        t.range.c0 <= shape.range.c1 &&
        shape.range.c0 <= t.range.c1,
    )
    return [
      {
        id: "table",
        label: "Table",
        tip: "Tables help you sort, filter, and summarize data.",
        ...(inTable ? { disabled: "The data is already in a table." } : {}),
      },
      // Excel's first recommended PivotTables, then More for the rest (Recommended PivotTables).
      ...(this.pivots(shape)?.items.slice(0, 3) ?? []).map((item, i) => ({
        id: `pivot:${i}`,
        label: "PivotTable",
        tip: `${item.recipe.title}. PivotTables summarize data, on a new sheet.`,
      })),
      { id: "more:pivot", label: "More", tip: "Click for more PivotTables." },
    ]
  }

  /** Quick Analysis stops suggesting PivotTables above this many rows (Recommended PivotTables still does). */
  private static PIVOT_ROWS = 50_000
  private pivotsFor: { key: string; found: ReturnType<SheetController["pivots"]["recommendations"]> } | null = null

  /** The recommended PivotTables for the selection (kept while the workbook doesn't change). */
  pivots(shape: QuickShape) {
    if (shape.range.r1 - shape.range.r0 > QuickAnalysis.PIVOT_ROWS) return null
    const source = rangeName(shape.range)
    const key = `${this.c.sheetIndex}:${source}:${this.c.wb.version}`
    if (this.pivotsFor?.key !== key) this.pivotsFor = { key, found: this.c.pivots.recommendations(source) }
    const found = this.pivotsFor.found
    return typeof found === "string" ? null : found
  }

  /** The data and cells of sparklines for each row (the numbers' columns; the sparklines in the column right). */
  private sparkRanges(shape: QuickShape): { data: Range; location: Range } | null {
    const cols = shape.numberCols
    if (cols.length < 2) return null
    const { range, top } = shape
    return {
      data: { r0: top, r1: range.r1, c0: cols[0], c1: cols[cols.length - 1] },
      location: { r0: top, r1: range.r1, c0: range.c1 + 1, c1: range.c1 + 1 },
    }
  }

  sparklines(shape: QuickShape): QuickItem[] {
    const none = this.sparkRanges(shape) ? {} : { disabled: "Sparklines need two or more columns of numbers." }
    return [
      { id: "spark:line", label: "Line", tip: "A line chart in each row's cell.", ...none },
      { id: "spark:column", label: "Column", tip: "A column chart in each row's cell.", ...none },
      { id: "spark:stacked", label: "Win/Loss", tip: "A win/loss chart in each row's cell.", ...none },
    ]
  }

  /** The changes a choice makes to cells (totals, a table, sparklines); null for the others. */
  private opsFor(id: string, shape: QuickShape): Op[] | null {
    const [what, a, b] = id.split(":")
    if (what === "total") return this.totalOps(shape, a as TotalSide, b as TotalKind).ops
    if (what === "table") {
      const ops = this.c.tables.createOps(shape.range, shape.header, DEFAULT_TABLE_STYLE)
      return typeof ops === "string" ? null : ops
    }
    if (what === "spark") {
      const at = this.sparkRanges(shape)
      if (!at) return null
      const ops = this.c.sparklines.createOps(a as SparklineType, at.data, at.location)
      return typeof ops === "string" ? null : ops
    }
    return null
  }

  // ── Preview and apply ──

  /** Shows a choice on the sheet while the pointer is over it (charts preview in the gallery instead). */
  preview(id: string) {
    this.endPreview()
    if (this.c.edit || this.c.sheet.protection) return
    if (id.startsWith("cf:")) {
      const rule = this.rule(id)
      if (rule) this.c.conditional.setPreview(rule)
      return
    }
    const shape = this.shape()
    const ops = shape && this.opsFor(id, shape)
    if (!ops) return
    previewing.add(this.c.wb)
    try {
      this.shown = this.c.wb.transact(ops, { label: "Preview", history: false }) as UndoEntry<unknown> | null
    } finally {
      previewing.delete(this.c.wb)
    }
  }

  /** Takes the preview back (the pointer left, the gallery closed). */
  endPreview() {
    if (this.c.conditional.preview) this.c.conditional.setPreview(null)
    const shown = this.shown
    if (!shown) return
    this.shown = null
    previewing.add(this.c.wb)
    try {
      this.c.wb.transact(shown.undo, { label: "Preview", history: false })
    } finally {
      previewing.delete(this.c.wb)
    }
  }

  get previewShown(): boolean {
    return this.shown !== null
  }

  /** A click on a choice: the preview goes, then the choice is made once (one undo step). */
  apply(id: string) {
    this.endPreview()
    const shape = this.shape()
    if (!shape) return
    const [what, a, b] = id.split(":")
    const sel = this.c.selection
    if (what === "cf") {
      if (a === "clear") return this.c.conditional.clearSelected()
      const rule = this.rule(id)
      if (rule && a !== "greaterThan" && a !== "lessThan") return this.c.conditional.add(rule)
      // Greater Than, Less Than, Equal To and Text Contains ask for the value, as Excel's do.
      return this.c.shell.openDialog?.(`cf:${a}`)
    }
    if (what === "chart") return void this.c.charts.insert(a)
    if (what === "more" && a === "pivot") return this.c.shell.openDialog?.(`recommendPivots:${rangeName(shape.range)}`)
    if (what === "more") return this.c.shell.showTab?.("TabInsert")
    if (what === "pivot") {
      const item = this.pivots(shape)?.items[Number(a)]
      if (!item) return
      const problem = this.c.pivots.create(rangeName(shape.range), { kind: "newSheet" }, undefined, item.recipe)
      if (problem) this.c.shell.alert(problem)
      return
    }
    if (what === "table") return void this.c.tables.create(shape.range, shape.header, DEFAULT_TABLE_STYLE)
    if (what === "total") {
      const { ops, cells } = this.totalOps(shape, a as TotalSide, b as TotalKind)
      const taken = cells.some((w) => {
        const d = this.c.sheet.cells.get(w.row, w.col)
        return d !== undefined && ((d.value !== undefined && d.value !== "") || d.input !== undefined)
      })
      const write = () => void this.c.transact(ops, { label: "Quick Analysis", before: sel, after: sel })
      if (taken) this.c.shell.confirm(REPLACE_QUESTION, write)
      else write()
      return
    }
    if (what === "spark") {
      const ops = this.opsFor(id, shape)
      if (!ops) return
      if (this.c.transact(ops, { label: "Insert Sparklines", before: sel, after: sel }))
        this.c.shell.showTab?.("TabSparkline")
    }
  }
}
