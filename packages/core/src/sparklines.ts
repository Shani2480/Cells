import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import { chartRefRange } from "./chart"
import { quoteSheet } from "./formula/print"
import type { Sheet } from "./sheet"
import { themeColorHex, type Theme, type ThemeColorRef } from "./theme"
import { columnName } from "./address"

/**
 * Sparklines (Insert › Sparklines), as Excel 2010+ keeps them: groups of tiny charts drawn inside cells, each group
 * with one type, one set of colours and options, and one sparkline per cell, each reading its own range of data.
 */
export type SparklineType = "line" | "column" | "stacked"

/** A colour: "#RRGGBB", or a theme colour (follows the workbook's theme). */
export type SparkColor = string | ThemeColorRef

export type SparkColors = {
  series: SparkColor
  negative: SparkColor
  axis: SparkColor
  markers: SparkColor
  first: SparkColor
  last: SparkColor
  high: SparkColor
  low: SparkColor
}

/** Vertical axis bounds: each sparkline its own, the same for the whole group, or a set number. */
export type SparkAxisScale = "individual" | "group" | "custom"

export type Sparkline = {
  /** The cell it is drawn in, on the group's sheet. */
  row: number
  col: number
  /** Its data: a range with the sheet's name ("Sheet1!A2:E2"); empty when none. */
  source: string
}

export type SparklineGroup = {
  type: SparklineType
  colors: SparkColors
  /** Line width in points (Excel's default 0.75). */
  lineWeight: number
  markers: boolean
  high: boolean
  low: boolean
  first: boolean
  last: boolean
  negative: boolean
  /** Show Axis: a horizontal line at zero. */
  displayXAxis: boolean
  /** Hidden & Empty Cell Settings: empty cells as gaps, zero, or (lines) joined over them. */
  emptyAs: "gap" | "zero" | "span"
  /** Show data in hidden rows and columns. */
  showHidden: boolean
  minAxis: SparkAxisScale
  maxAxis: SparkAxisScale
  manualMin?: number
  manualMax?: number
  /** Plot Data Right-to-Left. */
  rightToLeft: boolean
  /** Date Axis Type: the range of dates that spaces the points (with the sheet's name). */
  dateRange?: string
  sparklines: Sparkline[]
  /** Excel's revision id (xr2:uid), kept as read. */
  uid?: string
}

// ── Styles ────────────────────────────────────────────────────────────────────────────────────────────────────

type StyleColors = Omit<SparkColors, "axis">
const t = (theme: number, tint?: number): ThemeColorRef => (tint ? { theme, tint } : { theme })

/**
 * Excel's 36 Sparkline Styles in gallery order (six accents across, six rows): Accent N Darker 50%, Darker 25%,
 * Accent N, Lighter 40%, then Dark #1–6 and Colorful #1–6. Colours as Excel writes them.
 */
// prettier-ignore
export const SPARKLINE_STYLES: { name: string; colors: StyleColors }[] = [
  { name: "Accent 1, Darker 50%", colors: { series: t(4, -0.499984740745262), negative: t(5), markers: t(4, -0.499984740745262), high: t(4), low: t(4), first: t(4, 0.3999755851924192), last: t(4, 0.3999755851924192) } },
  { name: "Accent 2, Darker 50%", colors: { series: t(5, -0.499984740745262), negative: t(6), markers: t(5, -0.499984740745262), high: t(5), low: t(5), first: t(5, 0.3999755851924192), last: t(5, 0.3999755851924192) } },
  { name: "Accent 3, Darker 50%", colors: { series: t(6, -0.499984740745262), negative: t(7), markers: t(6, -0.499984740745262), high: t(6), low: t(6), first: t(6, 0.3999755851924192), last: t(6, 0.3999755851924192) } },
  { name: "Accent 4, Darker 50%", colors: { series: t(7, -0.499984740745262), negative: t(8), markers: t(7, -0.499984740745262), high: t(7), low: t(7), first: t(7, 0.3999755851924192), last: t(7, 0.3999755851924192) } },
  { name: "Accent 5, Darker 50%", colors: { series: t(8, -0.499984740745262), negative: t(9), markers: t(8, -0.499984740745262), high: t(8), low: t(8), first: t(8, 0.3999755851924192), last: t(8, 0.3999755851924192) } },
  { name: "Accent 6, Darker 50%", colors: { series: t(9, -0.499984740745262), negative: t(4), markers: t(9, -0.499984740745262), high: t(9), low: t(9), first: t(9, 0.3999755851924192), last: t(9, 0.3999755851924192) } },
  { name: "Accent 1, Darker 25%", colors: { series: t(4, -0.249977111117893), negative: t(5), markers: t(5, -0.249977111117893), high: t(5, -0.249977111117893), low: t(5, -0.249977111117893), first: t(5, -0.249977111117893), last: t(5, -0.249977111117893) } },
  { name: "Accent 2, Darker 25%", colors: { series: t(5, -0.249977111117893), negative: t(6), markers: t(6, -0.249977111117893), high: t(6, -0.249977111117893), low: t(6, -0.249977111117893), first: t(6, -0.249977111117893), last: t(6, -0.249977111117893) } },
  { name: "Accent 3, Darker 25%", colors: { series: t(6, -0.249977111117893), negative: t(7), markers: t(7, -0.249977111117893), high: t(7, -0.249977111117893), low: t(7, -0.249977111117893), first: t(7, -0.249977111117893), last: t(7, -0.249977111117893) } },
  { name: "Accent 4, Darker 25%", colors: { series: t(7, -0.249977111117893), negative: t(8), markers: t(8, -0.249977111117893), high: t(8, -0.249977111117893), low: t(8, -0.249977111117893), first: t(8, -0.249977111117893), last: t(8, -0.249977111117893) } },
  { name: "Accent 5, Darker 25%", colors: { series: t(8, -0.249977111117893), negative: t(9), markers: t(9, -0.249977111117893), high: t(9, -0.249977111117893), low: t(9, -0.249977111117893), first: t(9, -0.249977111117893), last: t(9, -0.249977111117893) } },
  { name: "Accent 6, Darker 25%", colors: { series: t(9, -0.249977111117893), negative: t(4), markers: t(4, -0.249977111117893), high: t(4, -0.249977111117893), low: t(4, -0.249977111117893), first: t(4, -0.249977111117893), last: t(4, -0.249977111117893) } },
  { name: "Accent 1", colors: { series: t(4), negative: t(5), markers: t(4, -0.249977111117893), high: t(4, -0.249977111117893), low: t(4, -0.249977111117893), first: t(4, -0.249977111117893), last: t(4, -0.249977111117893) } },
  { name: "Accent 2", colors: { series: t(5), negative: t(6), markers: t(5, -0.249977111117893), high: t(5, -0.249977111117893), low: t(5, -0.249977111117893), first: t(5, -0.249977111117893), last: t(5, -0.249977111117893) } },
  { name: "Accent 3", colors: { series: t(6), negative: t(7), markers: t(6, -0.249977111117893), high: t(6, -0.249977111117893), low: t(6, -0.249977111117893), first: t(6, -0.249977111117893), last: t(6, -0.249977111117893) } },
  { name: "Accent 4", colors: { series: t(7), negative: t(8), markers: t(7, -0.249977111117893), high: t(7, -0.249977111117893), low: t(7, -0.249977111117893), first: t(7, -0.249977111117893), last: t(7, -0.249977111117893) } },
  { name: "Accent 5", colors: { series: t(8), negative: t(9), markers: t(8, -0.249977111117893), high: t(8, -0.249977111117893), low: t(8, -0.249977111117893), first: t(8, -0.249977111117893), last: t(8, -0.249977111117893) } },
  { name: "Accent 6", colors: { series: t(9), negative: t(4), markers: t(9, -0.249977111117893), high: t(9, -0.249977111117893), low: t(9, -0.249977111117893), first: t(9, -0.249977111117893), last: t(9, -0.249977111117893) } },
  { name: "Accent 1, Lighter 40%", colors: { series: t(4, 0.3999755851924192), negative: t(0, -0.499984740745262), markers: t(4, 0.7999816888943144), high: t(4, -0.499984740745262), low: t(4, -0.499984740745262), first: t(4, -0.249977111117893), last: t(4, -0.249977111117893) } },
  { name: "Accent 2, Lighter 40%", colors: { series: t(5, 0.3999755851924192), negative: t(0, -0.499984740745262), markers: t(5, 0.7999816888943144), high: t(5, -0.499984740745262), low: t(5, -0.499984740745262), first: t(5, -0.249977111117893), last: t(5, -0.249977111117893) } },
  { name: "Accent 3, Lighter 40%", colors: { series: t(6, 0.3999755851924192), negative: t(0, -0.499984740745262), markers: t(6, 0.7999816888943144), high: t(6, -0.499984740745262), low: t(6, -0.499984740745262), first: t(6, -0.249977111117893), last: t(6, -0.249977111117893) } },
  { name: "Accent 4, Lighter 40%", colors: { series: t(7, 0.3999755851924192), negative: t(0, -0.499984740745262), markers: t(7, 0.7999816888943144), high: t(7, -0.499984740745262), low: t(7, -0.499984740745262), first: t(7, -0.249977111117893), last: t(7, -0.249977111117893) } },
  { name: "Accent 5, Lighter 40%", colors: { series: t(8, 0.3999755851924192), negative: t(0, -0.499984740745262), markers: t(8, 0.7999816888943144), high: t(8, -0.499984740745262), low: t(8, -0.499984740745262), first: t(8, -0.249977111117893), last: t(8, -0.249977111117893) } },
  { name: "Accent 6, Lighter 40%", colors: { series: t(9, 0.3999755851924192), negative: t(0, -0.499984740745262), markers: t(9, 0.7999816888943144), high: t(9, -0.499984740745262), low: t(9, -0.499984740745262), first: t(9, -0.249977111117893), last: t(9, -0.249977111117893) } },
  { name: "Dark #1", colors: { series: t(1, 0.499984740745262), negative: t(1, 0.249977111117893), markers: t(1, 0.249977111117893), high: t(1, 0.249977111117893), low: t(1, 0.249977111117893), first: t(1, 0.249977111117893), last: t(1, 0.249977111117893) } },
  { name: "Dark #2", colors: { series: t(1, 0.3499862666707358), negative: t(0, -0.249977111117893), markers: t(0, -0.249977111117893), high: t(0, -0.249977111117893), low: t(0, -0.249977111117893), first: t(0, -0.249977111117893), last: t(0, -0.249977111117893) } },
  { name: "Dark #3", colors: { series: "#323232", negative: "#D00000", markers: "#D00000", high: "#D00000", low: "#D00000", first: "#D00000", last: "#D00000" } },
  { name: "Dark #4", colors: { series: "#000000", negative: "#0070C0", markers: "#0070C0", high: "#0070C0", low: "#0070C0", first: "#0070C0", last: "#0070C0" } },
  { name: "Dark #5", colors: { series: "#376092", negative: "#D00000", markers: "#D00000", high: "#D00000", low: "#D00000", first: "#D00000", last: "#D00000" } },
  { name: "Dark #6", colors: { series: "#0070C0", negative: "#000000", markers: "#000000", high: "#000000", low: "#000000", first: "#000000", last: "#000000" } },
  { name: "Colorful #1", colors: { series: "#5F5F5F", negative: "#FFB620", markers: "#D70077", high: "#56BE79", low: "#FF5055", first: "#5687C2", last: "#359CEB" } },
  { name: "Colorful #2", colors: { series: "#5687C2", negative: "#FFB620", markers: "#D70077", high: "#56BE79", low: "#FF5055", first: "#777777", last: "#359CEB" } },
  { name: "Colorful #3", colors: { series: "#C6EFCE", negative: "#FFC7CE", markers: "#8CADD6", high: "#60D276", low: "#FF5367", first: "#FFDC47", last: "#FFEB9C" } },
  { name: "Colorful #4", colors: { series: "#00B050", negative: "#FF0000", markers: "#0070C0", high: "#00B050", low: "#FF0000", first: "#FFC000", last: "#FFC000" } },
  { name: "Colorful #5", colors: { series: t(3), negative: t(9), markers: t(8), high: t(6), low: t(7), first: t(4), last: t(5) } },
  { name: "Colorful #6", colors: { series: t(1), negative: t(9), markers: t(8), high: t(6), low: t(7), first: t(4), last: t(5) } },
]

/** The group colours of a style (Excel leaves the axis black). */
export function sparkStyleColors(index: number): SparkColors {
  return { ...(SPARKLINE_STYLES[index] ?? SPARKLINE_STYLES[0]).colors, axis: "#000000" }
}

/** The style a group's colours are, or -1 (colours changed by hand). */
export function sparkStyleIndex(colors: SparkColors): number {
  const key = (c: SparkColor) => (typeof c === "string" ? c.toUpperCase() : `${c.theme}:${(c.tint ?? 0).toFixed(4)}`)
  return SPARKLINE_STYLES.findIndex((s) =>
    (Object.keys(s.colors) as (keyof StyleColors)[]).every((k) => key(s.colors[k]) === key(colors[k])),
  )
}

/** A new group as Excel's Insert › Sparklines makes it (Accent 1, Darker 50%). */
export function newSparklineGroup(type: SparklineType, sparklines: Sparkline[]): SparklineGroup {
  return {
    type,
    colors: sparkStyleColors(0),
    lineWeight: 0.75,
    markers: false,
    high: false,
    low: false,
    first: false,
    last: false,
    negative: false,
    displayXAxis: false,
    emptyAs: "gap",
    showHidden: false,
    minAxis: "individual",
    maxAxis: "individual",
    rightToLeft: false,
    sparklines,
  }
}

export const sparkColorHex = (c: SparkColor, theme: Theme): string =>
  typeof c === "string" ? c : themeColorHex(c, theme)

export function copySparklineGroups(groups: readonly SparklineGroup[]): SparklineGroup[] {
  return groups.map((g) => ({ ...g, colors: { ...g.colors }, sparklines: g.sparklines.map((s) => ({ ...s })) }))
}

/** The group and sparkline drawn in a cell, or null. */
export function sparklineAt(
  groups: readonly SparklineGroup[],
  row: number,
  col: number,
): { group: SparklineGroup; sparkline: Sparkline } | null {
  for (const group of groups)
    for (const sparkline of group.sparklines)
      if (sparkline.row === row && sparkline.col === col) return { group, sparkline }
  return null
}

/** Sparklines follow inserted and deleted rows and columns; those whose cell is deleted go, then empty groups. */
export function shiftSparklines(
  groups: readonly SparklineGroup[],
  axis: "rows" | "cols",
  at: number,
  count: number,
): SparklineGroup[] {
  const out: SparklineGroup[] = []
  for (const g of groups) {
    const sparklines: Sparkline[] = []
    for (const s of g.sparklines) {
      const index = axis === "rows" ? s.row : s.col
      let moved = index
      if (count > 0) moved = index >= at ? index + count : index
      else if (index >= at && index < at - count) continue
      else if (index >= at - count) moved = index + count
      if (moved < 0 || moved >= (axis === "rows" ? MAX_ROWS : MAX_COLS)) continue
      sparklines.push(axis === "rows" ? { ...s, row: moved } : { ...s, col: moved })
    }
    if (sparklines.length) out.push({ ...g, colors: { ...g.colors }, sparklines })
  }
  return out
}

/** Rewrites every sparkline's data and date range; true when one changed. */
export function rewriteSparklineRefs(groups: SparklineGroup[], map: (ref: string) => string | null): boolean {
  let changed = false
  const go = (ref: string) => {
    if (!ref) return ref
    const next = map(ref)
    if (next === null || next === ref) return ref
    changed = true
    return next
  }
  for (const g of groups) {
    if (g.dateRange) g.dateRange = go(g.dateRange)
    for (const s of g.sparklines) s.source = go(s.source)
  }
  return changed
}

/** "Sheet1!A2:E2", as Excel writes a sparkline's data. */
export function sparkRefText(sheetName: string, range: Range): string {
  const cell = (row: number, col: number) => `${columnName(col)}${row + 1}`
  const one = range.r0 === range.r1 && range.c0 === range.c1
  return `${quoteSheet(sheetName)}!${cell(range.r0, range.c0)}${one ? "" : ":" + cell(range.r1, range.c1)}`
}

/**
 * Create Sparklines: one sparkline per cell of the location range (one row or one column), each reading the line of
 * the data range that matches it (rows for a column of locations, columns for a row). The error Excel shows when
 * the ranges don't fit, or the sparklines.
 */
export function sparklinesFor(
  dataSheet: string,
  data: Range,
  location: Range,
): { sparklines: Sparkline[] } | { error: string } {
  const lr = location.r1 - location.r0 + 1
  const lc = location.c1 - location.c0 + 1
  if (lr > 1 && lc > 1)
    return {
      error:
        "Location reference is not valid because the cells are not all in the same column or same row. Select cells that are all in a single row or column.",
    }
  const dr = data.r1 - data.r0 + 1
  const dc = data.c1 - data.c0 + 1
  const count = Math.max(lr, lc)
  const sparklines: Sparkline[] = []
  const at = (i: number) =>
    lr > 1 ? { row: location.r0 + i, col: location.c0 } : { row: location.r0, col: location.c0 + i }
  if (count === 1) {
    if (dr > 1 && dc > 1)
      return { error: "Data reference is not valid because the cells are not all in the same column or same row." }
    return { sparklines: [{ ...at(0), source: sparkRefText(dataSheet, data) }] }
  }
  // A column of locations takes the data's rows (or, if the data has as many columns, its columns), and so on.
  const byRows = lr > 1 ? dr === count : !(dc === count)
  if (byRows ? dr !== count : dc !== count)
    return {
      error: "The data reference is not valid because the number of rows or columns doesn't match the location range.",
    }
  for (let i = 0; i < count; i++) {
    const line = byRows ? { ...data, r0: data.r0 + i, r1: data.r0 + i } : { ...data, c0: data.c0 + i, c1: data.c0 + i }
    sparklines.push({ ...at(i), source: sparkRefText(dataSheet, line) })
  }
  return { sparklines }
}

// ── Reading the data ─────────────────────────────────────────────────────────────────────────────────────────

/**
 * A sparkline's points in order: numbers, null for empty cells (as the group's emptyAs says: kept as gaps, made 0,
 * or dropped for lines joined over them). Text counts as 0, as in charts. Hidden rows and columns are left out
 * unless the group shows them. `x` is each point's place (0…1) along the cell; with a date axis, by date.
 */
export function sparklinePoints(
  sheets: readonly Sheet[],
  group: SparklineGroup,
  sparkline: Sparkline,
): { values: (number | null)[]; x: number[] } {
  const at = sparkline.source ? chartRefRange(sheets, sparkline.source) : null
  if (!at) return { values: [], x: [] }
  const cells = lineCells(at.sheet, at.range, group.showHidden)
  let values: (number | null)[] = cells.map(([r, c]) => {
    const v = at.sheet.cells.get(r, c)?.value
    if (v === undefined || v === "") return group.emptyAs === "zero" ? 0 : null
    return typeof v === "number" ? v : typeof v === "boolean" ? null : typeof v === "string" ? 0 : null
  })
  let x = values.map((_, i) => (values.length > 1 ? i / (values.length - 1) : 0.5))
  if (group.dateRange) {
    const d = chartRefRange(sheets, group.dateRange)
    if (d) {
      const dates = lineCells(d.sheet, d.range, group.showHidden).map(([r, c]) => d.sheet.cells.get(r, c)?.value)
      const keep = values.map((_, i) => typeof dates[i] === "number")
      const numbers = dates.filter((v): v is number => typeof v === "number")
      const lo = Math.min(...numbers)
      const hi = Math.max(...numbers)
      if (numbers.length && hi > lo) {
        x = values.map((_, i) => (keep[i] ? ((dates[i] as number) - lo) / (hi - lo) : 0))
        values = values.map((v, i) => (keep[i] ? v : null))
      }
    }
  }
  if (group.rightToLeft) x = x.map((p) => 1 - p)
  return { values, x }
}

function lineCells(sheet: Sheet, range: Range, showHidden: boolean): [number, number][] {
  const out: [number, number][] = []
  for (let r = range.r0; r <= range.r1; r++)
    for (let c = range.c0; c <= range.c1; c++) {
      if (!showHidden && (sheet.rows.isHidden(r) || sheet.cols.isHidden(c))) continue
      out.push([r, c])
    }
  return out
}

/** The vertical bounds of each sparkline of a group, as its axis options say (null where it has no numbers). */
export function sparklineBounds(
  group: SparklineGroup,
  points: { values: (number | null)[] }[],
): ({ min: number; max: number } | null)[] {
  const own = points.map((p) => {
    const n = p.values.filter((v): v is number => v !== null)
    return n.length ? { min: Math.min(...n), max: Math.max(...n) } : null
  })
  const all = own.filter((b): b is { min: number; max: number } => b !== null)
  const groupMin = all.length ? Math.min(...all.map((b) => b.min)) : 0
  const groupMax = all.length ? Math.max(...all.map((b) => b.max)) : 0
  return own.map((b) => {
    if (!b) return null
    const min = group.minAxis === "custom" ? (group.manualMin ?? 0) : group.minAxis === "group" ? groupMin : b.min
    const max = group.maxAxis === "custom" ? (group.manualMax ?? 0) : group.maxAxis === "group" ? groupMax : b.max
    return { min, max }
  })
}
