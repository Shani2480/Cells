import { allTables } from "./table/table"
import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import { copyFormula } from "./formula/adjust"
import type { Node } from "./formula/ast"
import { compareScalars, compareText, toNumber } from "./formula/coerce"
import type { EvalContext } from "./formula/context"
import { evaluateFormula } from "./formula/evaluate"
import { parseFormula } from "./formula/parser"
import { isMatrix, isRef, type Area, type Scalar } from "./formula/types"
import type { Sheet } from "./sheet"
import type { Border, CellStyle } from "./styles"
import { isError, type CellValue } from "./values"

/**
 * Conditional formatting, as Excel models it: rules over ranges, in priority order (first wins where two set the
 * same thing). A rule either formats the cells it matches (font, fill, border, number format) or draws a data bar,
 * a colour scale or an icon from the cell's place among the range's numbers.
 */

/** What a matching rule changes (Excel's "differential format": only these fields). */
export type CFStyle = Pick<
  CellStyle,
  "numberFormat" | "bold" | "italic" | "underline" | "strike" | "fontColor" | "fill"
> & { borderTop?: Border; borderBottom?: Border; borderLeft?: Border; borderRight?: Border }

export type CellIsOperator =
  | "greaterThan"
  | "lessThan"
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThanOrEqual"
  | "lessThanOrEqual"

export type TimePeriod =
  | "yesterday"
  | "today"
  | "tomorrow"
  | "last7Days"
  | "lastWeek"
  | "thisWeek"
  | "nextWeek"
  | "lastMonth"
  | "thisMonth"
  | "nextMonth"

/** A threshold of a scale, bar or icon set ("cfvo"). `value` is a number or a formula (without "="). */
export type CFThreshold = {
  type: "min" | "max" | "num" | "percent" | "percentile" | "formula" | "automatic"
  value?: string
}

export type IconSetName =
  | "3Arrows"
  | "3ArrowsGray"
  | "4Arrows"
  | "5Arrows"
  | "3TrafficLights1"
  | "3Signs"
  | "3Symbols"
  | "3Flags"
  | "3Stars"
  | "5Quarters"

export type CFRule =
  /** "Format only cells with: Cell Value …". Values are formulas without "=" ("5", "$A$1", "\"x\""). */
  | { kind: "cellIs"; operator: CellIsOperator; values: string[]; style: CFStyle }
  | { kind: "text"; operator: "contains" | "notContains" | "beginsWith" | "endsWith"; text: string; style: CFStyle }
  | { kind: "date"; period: TimePeriod; style: CFStyle }
  | { kind: "duplicate"; unique: boolean; style: CFStyle }
  | { kind: "top"; bottom: boolean; percent: boolean; rank: number; style: CFStyle }
  | { kind: "average"; below: boolean; equal: boolean; style: CFStyle }
  | { kind: "blanks"; blank: boolean; style: CFStyle }
  | { kind: "errors"; errors: boolean; style: CFStyle }
  /** "Use a formula": written for the top-left cell of the first range, relative references follow each cell. */
  | { kind: "expression"; formula: string; style: CFStyle }
  | { kind: "colorScale"; points: { at: CFThreshold; color: string }[] }
  | { kind: "dataBar"; min: CFThreshold; max: CFThreshold; color: string; gradient: boolean }
  | { kind: "iconSet"; set: IconSetName; reverse: boolean; thresholds: CFThreshold[]; showValue: boolean }

export type CFEntry = { id: number; ranges: Range[]; rule: CFRule; stopIfTrue: boolean }

/** What conditional formatting does to one cell. */
export type CFResult = {
  style?: CFStyle
  /** Colour scale fill (over any fill). */
  fill?: string
  /** Data bar: from `start` to `end` (shares of the cell width); negative bars go left of the axis. */
  bar?: { start: number; end: number; axis: number | null; color: string; gradient: boolean; negative: boolean }
  icon?: { set: IconSetName; index: number; count: number }
  /** The icon set hides the number. */
  hideValue?: boolean
}

/** One icon of an icon set (as Filter by Cell Icon and Sort On Conditional Formatting Icon name it). */
export type IconRef = { set: IconSetName; index: number }

/** What conditional formatting does to a cell of the sheet being filtered or sorted (null: nothing). */
export type CellLook = (row: number, col: number) => CFResult | null

/**
 * The fill and font colour a cell shows, conditional formats included (Filter and Sort by Color list and match
 * these, as Excel does): a colour scale over a rule's fill over the cell's own.
 */
export function shownColors(style: CellStyle, cf: CFResult | null): { fill: string | null; font: string | null } {
  return {
    fill: (cf?.fill ?? cf?.style?.fill ?? style.fill ?? null)?.toUpperCase() ?? null,
    font: (cf?.style?.fontColor ?? style.fontColor ?? null)?.toUpperCase() ?? null,
  }
}

/** The icon a cell shows, or null. */
export function shownIcon(cf: CFResult | null): IconRef | null {
  return cf?.icon ? { set: cf.icon.set, index: cf.icon.index } : null
}

export const sameIcon = (a: IconRef | null | undefined, b: IconRef | null | undefined) =>
  (a ?? null) === (b ?? null) || (!!a && !!b && a.set === b.set && a.index === b.index)

// ── Editing ───────────────────────────────────────────────────────────────────────────────────────────────────

/** Moves a range for inserted (count > 0) or deleted (count < 0) lines; null when it is deleted whole. */
export function shiftRange(range: Range, axis: "rows" | "cols", at: number, count: number): Range | null {
  const lo = axis === "rows" ? "r0" : "c0"
  const hi = axis === "rows" ? "r1" : "c1"
  const limit = (axis === "rows" ? MAX_ROWS : MAX_COLS) - 1
  const next = { ...range }
  if (count > 0) {
    if (range[lo] >= at) next[lo] += count
    if (range[hi] >= at) next[hi] += count
    if (next[lo] > limit) return null
    next[hi] = Math.min(next[hi], limit)
    return next
  }
  const end = at - count - 1
  next[lo] = range[lo] < at ? range[lo] : range[lo] > end ? range[lo] + count : at
  next[hi] = range[hi] > end ? range[hi] + count : range[hi] >= at ? at - 1 : range[hi]
  return next[hi] < next[lo] ? null : next
}

/** Rules follow inserted and deleted rows and columns; a rule whose ranges all go is dropped. */
export function shiftConditional(sheet: Sheet, axis: "rows" | "cols", at: number, count: number) {
  if (!sheet.conditional.length) return
  sheet.conditional = sheet.conditional
    .map((e) => ({ ...e, ranges: e.ranges.map((r) => shiftRange(r, axis, at, count)).filter((r) => r !== null) }))
    .filter((e) => e.ranges.length > 0)
}

export function copyEntries(entries: readonly CFEntry[]): CFEntry[] {
  return entries.map((e) => JSON.parse(JSON.stringify(e)) as CFEntry)
}

const overlaps = (a: Range, b: Range) => a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/**
 * Clear Rules from Selected Cells: the cells of `ranges` leave every rule; a rule keeps the rest of its ranges
 * (split into rectangles around the cleared part).
 */
export function clearRulesIn(entries: readonly CFEntry[], ranges: readonly Range[]): CFEntry[] {
  const out: CFEntry[] = []
  for (const e of entries) {
    let parts = e.ranges.map((r) => ({ ...r }))
    for (const cut of ranges) parts = parts.flatMap((p) => subtract(p, cut))
    if (parts.length) out.push({ ...(JSON.parse(JSON.stringify(e)) as CFEntry), ranges: parts })
  }
  return out
}

/** `a` less `b`, as up to four rectangles (above, below, left, right). */
export function subtract(a: Range, b: Range): Range[] {
  if (!overlaps(a, b)) return [a]
  const out: Range[] = []
  if (a.r0 < b.r0) out.push({ ...a, r1: b.r0 - 1 })
  if (a.r1 > b.r1) out.push({ ...a, r0: b.r1 + 1 })
  const r0 = Math.max(a.r0, b.r0)
  const r1 = Math.min(a.r1, b.r1)
  if (a.c0 < b.c0) out.push({ r0, r1, c0: a.c0, c1: b.c0 - 1 })
  if (a.c1 > b.c1) out.push({ r0, r1, c0: b.c1 + 1, c1: a.c1 })
  return out
}

/** Rules that touch any of `ranges` (Manage Rules › Show formatting rules for: Current Selection). */
export function rulesIn(entries: readonly CFEntry[], ranges: readonly Range[]): CFEntry[] {
  return entries.filter((e) => e.ranges.some((r) => ranges.some((s) => overlaps(r, s))))
}

// ── Evaluation ────────────────────────────────────────────────────────────────────────────────────────────────

/** Reads the values the workbook holds (calculated elsewhere): conditions never calculate cells themselves. */
export class ValueContext implements EvalContext {
  sheet = 0
  row = 0
  col = 0
  constructor(
    private readonly sheets: readonly Sheet[],
    private readonly today: number,
  ) {}
  cell(sheet: number, row: number, col: number): Scalar {
    return this.sheets[sheet]?.cells.value(row, col) ?? null
  }
  values(area: Area, visit: (value: Scalar, row: number, col: number) => void) {
    const cells = this.sheets[area.sheet]?.cells
    cells?.forEach(area, (row, col, data) => data.value !== undefined && visit(data.value, row, col), true)
  }
  usedArea(area: Area): Area | null {
    const used = this.sheets[area.sheet]?.cells.usedBounds()
    if (!used) return null
    const r = {
      r0: Math.max(area.r0, used.r0),
      r1: Math.min(area.r1, used.r1),
      c0: Math.max(area.c0, used.c0),
      c1: Math.min(area.c1, used.c1),
    }
    return r.r0 > r.r1 || r.c0 > r.c1 ? null : { sheet: area.sheet, ...r }
  }
  sheetIndex(name: string) {
    const upper = name.toUpperCase()
    return this.sheets.findIndex((s) => s.name.toUpperCase() === upper)
  }
  sheetName(index: number) {
    return this.sheets[index]?.name ?? ""
  }
  sheetCount() {
    return this.sheets.length
  }
  tables() {
    return allTables(this.sheets)
  }
  defined(name: string) {
    return (this.sheets as { names?: { node(name: string): Node | null } }).names?.node(name) ?? null
  }
  spillArea() {
    return null
  }
  isFormula(sheet: number, row: number, col: number) {
    return this.sheets[sheet]?.cells.input(row, col)?.startsWith("=") ?? false
  }
  formulaText(sheet: number, row: number, col: number) {
    return this.sheets[sheet]?.cells.input(row, col) ?? null
  }
  now() {
    return this.today
  }
  random() {
    return Math.random()
  }
}

const isNumber = (v: CellValue | undefined): v is number => typeof v === "number"

/** Numbers in a rule's ranges (only the used part of whole rows and columns is read). */
function numbersIn(sheet: Sheet, ranges: readonly Range[]): number[] {
  const out: number[] = []
  for (const r of ranges)
    sheet.cells.forEach(r, (_row, _col, data) => isNumber(data.value) && out.push(data.value), true)
  return out
}

/** The k-th largest (1-based) of sorted-descending numbers. */
function percentileOf(sorted: number[], p: number): number {
  // Excel's PERCENTILE (inclusive) over the ascending values.
  const asc = sorted
  const n = asc.length
  if (!n) return 0
  const x = (p / 100) * (n - 1)
  const lo = Math.floor(x)
  const hi = Math.min(n - 1, lo + 1)
  return asc[lo] + (asc[hi] - asc[lo]) * (x - lo)
}

function mixColor(a: string, b: string, t: number): string {
  const pa = parseInt(a.slice(1), 16)
  const pb = parseInt(b.slice(1), 16)
  const ch = (p: number, s: number) => (p >> s) & 255
  const c = (s: number) => Math.round(ch(pa, s) + (ch(pb, s) - ch(pa, s)) * t)
  return "#" + ((c(16) << 16) | (c(8) << 8) | c(0)).toString(16).padStart(6, "0").toUpperCase()
}

/** Days since the epoch → day of week, Sunday = 0 (serial 1 = Sunday 1 Jan 1900). */
const weekday = (serial: number) => ((Math.floor(serial) % 7) + 6) % 7

function monthOf(serial: number): number {
  // Serials before 1 March 1900 are off by Excel's phantom day; dates that old are not worth the special case here.
  const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86_400_000)
  return d.getUTCFullYear() * 12 + d.getUTCMonth()
}

function inPeriod(serial: number, period: TimePeriod, today: number): boolean {
  const d = Math.floor(serial)
  const t = Math.floor(today)
  const weekStart = t - weekday(t)
  switch (period) {
    case "yesterday":
      return d === t - 1
    case "today":
      return d === t
    case "tomorrow":
      return d === t + 1
    case "last7Days":
      return d > t - 7 && d <= t
    case "lastWeek":
      return d >= weekStart - 7 && d < weekStart
    case "thisWeek":
      return d >= weekStart && d < weekStart + 7
    case "nextWeek":
      return d >= weekStart + 7 && d < weekStart + 14
    case "lastMonth":
      return monthOf(d) === monthOf(t) - 1
    case "thisMonth":
      return monthOf(d) === monthOf(t)
    case "nextMonth":
      return monthOf(d) === monthOf(t) + 1
  }
}

/** Icon counts and default thresholds (percent) per set. */
export const ICON_SETS: Record<IconSetName, number> = {
  "3Arrows": 3,
  "3ArrowsGray": 3,
  "4Arrows": 4,
  "5Arrows": 5,
  "3TrafficLights1": 3,
  "3Signs": 3,
  "3Symbols": 3,
  "3Flags": 3,
  "3Stars": 3,
  "5Quarters": 5,
}

export function defaultIconThresholds(count: number): CFThreshold[] {
  return Array.from({ length: count - 1 }, (_, i) => ({
    type: "percent",
    value: String(Math.round(((i + 1) * 100) / count)),
  }))
}

type Stats = { sorted: number[]; min: number; max: number; mean: number; counts?: Map<string, number> }

/**
 * Works out conditional formats for one sheet at one moment. Range-wide figures (top 10, averages, duplicates,
 * scales) are computed once per rule and kept while this object lives; make a new one when values change.
 */
export class ConditionalEval {
  private readonly ctx: ValueContext
  private readonly stats = new Map<number, Stats>()
  private readonly parsed = new Map<string, Node | null>()

  constructor(
    private readonly sheets: readonly Sheet[],
    private readonly sheetIndex: number,
    private readonly today: number,
    /** The rules to apply, when not the sheet's own (a dialog's preview). */
    private readonly entries?: readonly CFEntry[],
  ) {
    this.ctx = new ValueContext(sheets, today)
  }

  private get sheet() {
    return this.sheets[this.sheetIndex]
  }

  /** The combined result of every rule over the cell (null when none applies). */
  at(row: number, col: number): CFResult | null {
    const entries = this.entries ?? this.sheet.conditional
    if (!entries.length) return null
    let result: CFResult | null = null
    for (const e of entries) {
      const range = e.ranges.find((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)
      if (!range) continue
      const hit = this.apply(e, row, col)
      if (!hit) continue
      result ??= {}
      if (hit.style) {
        // Higher-priority rules come first; later ones only fill in what is still unset.
        result.style = { ...hit.style, ...result.style }
      }
      if (hit.fill && result.fill === undefined) result.fill = hit.fill
      if (hit.bar && !result.bar) result.bar = hit.bar
      if (hit.icon && !result.icon) {
        result.icon = hit.icon
        result.hideValue = hit.hideValue
      }
      if (e.stopIfTrue) break
    }
    return result
  }

  /** Whether a "format" rule matches the cell (for tests and Manage Rules previews). */
  matches(entry: CFEntry, row: number, col: number): boolean {
    return this.apply(entry, row, col) !== null
  }

  private statsOf(e: CFEntry): Stats {
    let s = this.stats.get(e.id)
    if (s) return s
    const sorted = numbersIn(this.sheet, e.ranges).sort((a, b) => a - b)
    const sum = sorted.reduce((a, b) => a + b, 0)
    s = {
      sorted,
      min: sorted[0] ?? 0,
      max: sorted[sorted.length - 1] ?? 0,
      mean: sorted.length ? sum / sorted.length : 0,
    }
    if (e.rule.kind === "duplicate") {
      const counts = new Map<string, number>()
      for (const r of e.ranges)
        this.sheet.cells.forEach(
          r,
          (_row, _col, data) => {
            const k = keyOf(data.value)
            if (k !== null) counts.set(k, (counts.get(k) ?? 0) + 1)
          },
          true,
        )
      s.counts = counts
    }
    this.stats.set(e.id, s)
    return s
  }

  /** A formula's value at a cell, its relative references moved from the rule's top-left cell. */
  private evalAt(formula: string, e: CFEntry, row: number, col: number): Scalar {
    const origin = e.ranges[0]
    const dr = row - origin.r0
    const dc = col - origin.c0
    const key = `${formula}\u0000${dr}\u0000${dc}`
    let node = this.parsed.get(key)
    if (node === undefined) {
      const text = copyFormula("=" + formula, dr, dc)
      const parsed = parseFormula(text.slice(1))
      node = parsed.ok ? parsed.ast : null
      if (this.parsed.size > 50_000) this.parsed.clear()
      this.parsed.set(key, node)
    }
    if (!node) return { error: "#NAME?" }
    this.ctx.sheet = this.sheetIndex
    this.ctx.row = row
    this.ctx.col = col
    const v = evaluateFormula(node, this.ctx)
    if (isRef(v) || isMatrix(v)) return null
    return v
  }

  private threshold(t: CFThreshold, s: Stats, e: CFEntry, fallback: number): number {
    const n = Number(t.value)
    switch (t.type) {
      case "min":
        return s.min
      case "max":
        return s.max
      case "automatic":
        return fallback
      case "num":
        return Number.isFinite(n) ? n : fallback
      case "percent":
        return s.min + ((s.max - s.min) * (Number.isFinite(n) ? n : 0)) / 100
      case "percentile":
        return percentileOf(s.sorted, Number.isFinite(n) ? n : 50)
      case "formula": {
        const v = this.evalAt(t.value ?? "0", e, e.ranges[0].r0, e.ranges[0].c0)
        const num = toNumber(v)
        return typeof num === "number" ? num : fallback
      }
    }
  }

  private apply(e: CFEntry, row: number, col: number): CFResult | null {
    const rule = e.rule
    const value = this.sheet.cells.value(row, col)
    switch (rule.kind) {
      case "cellIs": {
        const a = this.evalAt(rule.values[0] ?? "0", e, row, col)
        const b = rule.values[1] === undefined ? null : this.evalAt(rule.values[1], e, row, col)
        return cellIs(rule.operator, value, a, b) ? { style: rule.style } : null
      }
      case "text": {
        if (value === undefined || isError(value)) return rule.operator === "notContains" ? { style: rule.style } : null
        const t = String(value).toLowerCase()
        const q = rule.text.toLowerCase()
        const hit =
          rule.operator === "contains"
            ? t.includes(q)
            : rule.operator === "notContains"
              ? !t.includes(q)
              : rule.operator === "beginsWith"
                ? t.startsWith(q)
                : t.endsWith(q)
        return hit ? { style: rule.style } : null
      }
      case "date":
        return isNumber(value) && inPeriod(value, rule.period, this.today) ? { style: rule.style } : null
      case "duplicate": {
        const k = keyOf(value)
        if (k === null) return null
        const n = this.statsOf(e).counts?.get(k) ?? 0
        return (rule.unique ? n === 1 : n > 1) ? { style: rule.style } : null
      }
      case "top": {
        if (!isNumber(value)) return null
        const s = this.statsOf(e)
        const count = s.sorted.length
        if (!count) return null
        let k = rule.percent ? Math.floor((count * rule.rank) / 100) : rule.rank
        k = Math.max(1, Math.min(count, k))
        const cut = rule.bottom ? s.sorted[k - 1] : s.sorted[count - k]
        return (rule.bottom ? value <= cut : value >= cut) ? { style: rule.style } : null
      }
      case "average": {
        if (!isNumber(value)) return null
        const { mean } = this.statsOf(e)
        const hit = rule.below ? (rule.equal ? value <= mean : value < mean) : rule.equal ? value >= mean : value > mean
        return hit ? { style: rule.style } : null
      }
      case "blanks": {
        const blank = value === undefined || (typeof value === "string" && value.trim() === "")
        return blank === rule.blank ? { style: rule.style } : null
      }
      case "errors": {
        const err = value !== undefined && isError(value)
        return err === rule.errors ? { style: rule.style } : null
      }
      case "expression": {
        const v = this.evalAt(rule.formula, e, row, col)
        const truthy = typeof v === "boolean" ? v : typeof v === "number" ? v !== 0 : false
        return truthy ? { style: rule.style } : null
      }
      case "colorScale": {
        if (!isNumber(value)) return null
        const s = this.statsOf(e)
        const pts = rule.points.map((p, i) => ({
          at: this.threshold(p.at, s, e, i === 0 ? s.min : s.max),
          color: p.color,
        }))
        if (value <= pts[0].at) return { fill: pts[0].color }
        const last = pts[pts.length - 1]
        if (value >= last.at) return { fill: last.color }
        for (let i = 1; i < pts.length; i++) {
          if (value <= pts[i].at) {
            const span = pts[i].at - pts[i - 1].at
            return { fill: mixColor(pts[i - 1].color, pts[i].color, span ? (value - pts[i - 1].at) / span : 0) }
          }
        }
        return null
      }
      case "dataBar": {
        if (!isNumber(value)) return null
        const s = this.statsOf(e)
        // Automatic (Excel 2010): the bars start at zero unless every number is on one side of it.
        const lo = this.threshold(rule.min, s, e, Math.min(0, s.min))
        const hi = this.threshold(rule.max, s, e, Math.max(0, s.max))
        if (hi <= lo)
          return {
            bar: {
              start: 0,
              end: value > lo ? 1 : 0,
              axis: null,
              color: rule.color,
              gradient: rule.gradient,
              negative: false,
            },
          }
        const v = Math.min(hi, Math.max(lo, value))
        if (lo < 0 && hi > 0) {
          const axis = -lo / (hi - lo)
          const at = (v - lo) / (hi - lo)
          const negative = v < 0
          return {
            bar: {
              start: negative ? at : axis,
              end: negative ? axis : at,
              axis,
              color: negative ? "#FF0000" : rule.color,
              gradient: rule.gradient,
              negative,
            },
          }
        }
        return {
          bar: {
            start: 0,
            end: (v - lo) / (hi - lo),
            axis: null,
            color: rule.color,
            gradient: rule.gradient,
            negative: false,
          },
        }
      }
      case "iconSet": {
        if (!isNumber(value)) return null
        const s = this.statsOf(e)
        const count = ICON_SETS[rule.set]
        let index = 0
        rule.thresholds.forEach((t) => {
          if (value >= this.threshold(t, s, e, s.min)) index++
        })
        index = Math.min(count - 1, index)
        if (rule.reverse) index = count - 1 - index
        return { icon: { set: rule.set, index, count }, hideValue: !rule.showValue }
      }
    }
  }
}

/** Duplicate Values compares text without case, numbers by value; blanks don't count. */
function keyOf(value: CellValue | undefined): string | null {
  if (value === undefined || value === "") return null
  if (typeof value === "string") return "s" + value.toLowerCase()
  if (typeof value === "number") return "n" + value
  if (typeof value === "boolean") return "b" + value
  return "e" + value.error
}

function cellIs(op: CellIsOperator, value: CellValue | undefined, a: Scalar, b: Scalar | null): boolean {
  if (value !== undefined && isError(value)) return false
  if ((a !== null && typeof a === "object") || (b !== null && typeof b === "object")) return false
  const v: Scalar = value === undefined ? null : value
  const cmp = (x: Scalar) => compareScalars(v as Exclude<Scalar, object>, x as Exclude<Scalar, object>)
  switch (op) {
    case "greaterThan":
      return cmp(a) > 0
    case "lessThan":
      return cmp(a) < 0
    case "greaterThanOrEqual":
      return cmp(a) >= 0
    case "lessThanOrEqual":
      return cmp(a) <= 0
    case "equal":
      return typeof v === "string" && typeof a === "string" ? compareText(v, a) === 0 : cmp(a) === 0
    case "notEqual":
      return cmp(a) !== 0
    case "between":
    case "notBetween": {
      // Either order: "between 10 and 1" is the same as "between 1 and 10".
      const inside = b !== null && ((cmp(a) >= 0 && cmp(b) <= 0) || (cmp(b) >= 0 && cmp(a) <= 0))
      return op === "between" ? inside : !inside
    }
  }
}
