// Timelines (PivotTable Analyze › Insert Timeline, Insert › Timeline): a strip of years, quarters, months or days for
// a date field, filtering the PivotTables it is connected to. Excel keeps the pick as a whole-day date filter
// ("dateBetween", useWholeDay) on that field of each connected PivotTable, on an axis or not (io's
// fixtures/excel/pivot-timeline.xlsx), so a timeline's state is its PivotTables' filters. The strip spans whole years
// of the field's dates; the look follows Excel 365's pictures of TimeSlicerStyleLight1 and Dark1.
import { dateSerial, serialToDate } from "../formula/dates"
import { applyTint, type Theme } from "../theme"
import type { Box } from "../slicers"
import { fieldShared } from "./group"
import type { PivotCache, PivotFilter, PivotTable } from "./model"
import type { PivotSlicerTarget } from "./slicer"

export type TimelineLevel = "years" | "quarters" | "months" | "days"
/** Excel's numbers for the levels (the timeline part's `level` and `selectionLevel`). */
export const TIMELINE_LEVELS: readonly TimelineLevel[] = ["years", "quarters", "months", "days"]

export const TIMELINE_STYLES = [
  "TimeSlicerStyleLight1",
  "TimeSlicerStyleLight2",
  "TimeSlicerStyleLight3",
  "TimeSlicerStyleLight4",
  "TimeSlicerStyleLight5",
  "TimeSlicerStyleLight6",
  "TimeSlicerStyleDark1",
  "TimeSlicerStyleDark2",
  "TimeSlicerStyleDark3",
  "TimeSlicerStyleDark4",
  "TimeSlicerStyleDark5",
  "TimeSlicerStyleDark6",
] as const
export const DEFAULT_TIMELINE_STYLE = "TimeSlicerStyleLight1"
/** A new timeline's size (px at 100%), as Excel makes it (4.72 × 1.67 in). */
export const TIMELINE_SIZE = { w: 453, h: 160 }

const MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
const cap = (m: string) => m[0] + m.slice(1).toLowerCase()

/** The first day of the period `serial` is in. */
export function periodStart(serial: number, level: TimelineLevel): number {
  const d = serialToDate(serial)
  if (level === "years") return dateSerial(d.year, 1, 1)
  if (level === "quarters") return dateSerial(d.year, Math.floor((d.month - 1) / 3) * 3 + 1, 1)
  if (level === "months") return dateSerial(d.year, d.month, 1)
  return Math.floor(serial)
}

/** The first day of the next period. */
export function nextPeriod(start: number, level: TimelineLevel): number {
  const d = serialToDate(start)
  if (level === "years") return dateSerial(d.year + 1, 1, 1)
  if (level === "quarters") return dateSerial(d.year, d.month + 3, 1)
  if (level === "months") return dateSerial(d.year, d.month + 1, 1)
  return start + 1
}

/** The span a timeline shows: whole years around the field's dates (`end` is the first day after). */
export function timelineBounds(cache: PivotCache, field: number): { start: number; end: number } | null {
  let lo = Infinity
  let hi = -Infinity
  for (const v of fieldShared(cache, field))
    if (typeof v === "number") {
      lo = Math.min(lo, v)
      hi = Math.max(hi, v)
    }
  if (!Number.isFinite(lo)) return null
  return { start: periodStart(lo, "years"), end: nextPeriod(periodStart(hi, "years"), "years") }
}

export type TimelinePeriod = {
  start: number
  /** The first day after it. */
  end: number
  /** Under the year line: "2025", "Q1", "JAN", "13". */
  label: string
  /** Above, at the first period of each year (or month, for days): "2025", "JAN 2025". */
  group: string | null
}

/** The periods of a level across the bounds. */
export function timelinePeriods(bounds: { start: number; end: number }, level: TimelineLevel): TimelinePeriod[] {
  const out: TimelinePeriod[] = []
  let prevGroup = ""
  for (let s = periodStart(bounds.start, level); s < bounds.end && out.length < 100000; s = nextPeriod(s, level)) {
    const d = serialToDate(s)
    const label =
      level === "years"
        ? String(d.year)
        : level === "quarters"
          ? `Q${Math.floor((d.month - 1) / 3) + 1}`
          : level === "months"
            ? MONTHS[d.month - 1]
            : String(d.day)
    const group = level === "years" ? "" : level === "days" ? `${MONTHS[d.month - 1]} ${d.year}` : String(d.year)
    out.push({ start: s, end: nextPeriod(s, level), label, group: group && group !== prevGroup ? group : null })
    prevGroup = group
  }
  return out
}

/** The dates a timeline picks: its first PivotTable's date filter on the field (`to` is the last day), or null. */
export function timelineSelection(target: PivotSlicerTarget): { from: number; to: number } | null {
  const f = target.tables[0].table.filters?.find((x) => x.kind === "date" && x.field === target.field)
  return f && f.kind === "date" ? { from: f.value, to: Math.floor(f.value2) } : null
}

/** The table with a timeline's dates on field `f` (null: none, Clear Filter). The same object when nothing changes. */
export function withTimelineDates(
  table: PivotTable,
  f: number,
  range: { from: number; to: number } | null,
): PivotTable {
  const filters = table.filters ?? []
  const was = filters.find((x) => x.kind === "date" && x.field === f)
  if (!range && !was) return table
  if (range && was && was.kind === "date" && was.value === range.from && was.value2 === range.to) return table
  const rest = filters.filter((x) => x !== was)
  if (!range) {
    const next: PivotTable = { ...table, filters: rest }
    if (!rest.length) delete next.filters
    return next
  }
  const id = was?.id ?? Math.max(0, ...filters.map((x) => x.id)) + 1
  const filter: PivotFilter = { field: f, id, kind: "date", op: "between", value: range.from, value2: range.to }
  return { ...table, filters: [...rest, filter] }
}

/** The selection label: "All Periods", or the picked dates at the level they were picked ("Mar - Jun 2025"). */
export function timelineLabel(range: { from: number; to: number } | null, level: TimelineLevel): string {
  if (!range) return "All Periods"
  const a = serialToDate(range.from)
  const b = serialToDate(range.to)
  const sameYear = a.year === b.year
  if (level === "years") return a.year === b.year ? String(a.year) : `${a.year} - ${b.year}`
  if (level === "quarters") {
    const qa = `Q${Math.floor((a.month - 1) / 3) + 1}`
    const qb = `Q${Math.floor((b.month - 1) / 3) + 1}`
    if (sameYear) return qa === qb ? `${qa} ${a.year}` : `${qa} - ${qb} ${a.year}`
    return `${qa} ${a.year} - ${qb} ${b.year}`
  }
  const ma = cap(MONTHS[a.month - 1])
  const mb = cap(MONTHS[b.month - 1])
  if (level === "months") {
    if (sameYear) return a.month === b.month ? `${ma} ${a.year}` : `${ma} - ${mb} ${a.year}`
    return `${ma} ${a.year} - ${mb} ${b.year}`
  }
  if (range.from === range.to) return `${ma} ${a.day}, ${a.year}`
  if (sameYear) return `${ma} ${a.day} - ${mb} ${b.day}, ${a.year}`
  return `${ma} ${a.day}, ${a.year} - ${mb} ${b.day}, ${b.year}`
}

// ── Layout and look ────────────────────────────────────────────────────────────────────────────────────────────

/** Each period's width (px at 100%) at each level, as Excel draws them. */
export const TIMELINE_PERIOD_WIDTH: Record<TimelineLevel, number> = { years: 60, quarters: 52, months: 44, days: 30 }

export type TimelineShow = { header: boolean; selectionLabel: boolean; timeLevel: boolean; scrollbar: boolean }

export type TimelineLayout = {
  header: Box | null
  clear: Box | null
  divider: number | null
  /** The selection label's and the time level's row (null: both hidden). */
  info: Box | null
  level: Box | null
  /** Where the year labels, period labels and bars sit. */
  groupY: number
  labelY: number
  bars: Box
  /** Each shown period's bar, from `first` on. */
  periods: Box[]
  first: number
  scrollbar: { box: Box; left: Box; right: Box; thumb: Box } | null
}

/** Where things sit in a timeline `w` wide with `count` periods, scrolled to period `first`. */
export function timelineLayout(
  show: TimelineShow,
  level: TimelineLevel,
  w: number,
  count: number,
  first: number,
): TimelineLayout {
  const pad = 8
  let y = 1
  const header = show.header ? { x: 1, y: 1, w: w - 2, h: 26 } : null
  if (header) y = 28
  const info = show.selectionLabel || show.timeLevel ? { x: pad, y, w: w - 2 * pad, h: 24 } : null
  if (info) y += 24
  else y += 4
  const groupY = y + 10
  const labelY = y + 31
  const bars = { x: pad, y: y + 42, w: w - 2 * pad, h: 17 }
  const pw = TIMELINE_PERIOD_WIDTH[level]
  const fit = Math.max(1, Math.floor(bars.w / pw))
  const start = Math.max(0, Math.min(first, Math.max(0, count - fit)))
  const periods: Box[] = []
  for (let i = start; i < count && (i - start) * pw < bars.w; i++)
    periods.push({
      x: bars.x + (i - start) * pw,
      y: bars.y,
      w: Math.min(pw, bars.x + bars.w - (bars.x + (i - start) * pw)),
      h: bars.h,
    })
  let scrollbar: TimelineLayout["scrollbar"] = null
  if (show.scrollbar) {
    const box = { x: pad, y: bars.y + bars.h + 7, w: w - 2 * pad, h: 16 }
    const track = box.w - 32
    const thumbW = Math.max(12, count ? (track * Math.min(fit, count)) / count : track)
    const room = track - thumbW
    const at = count > fit ? (room * start) / (count - fit) : 0
    scrollbar = {
      box,
      left: { x: box.x, y: box.y, w: 16, h: 16 },
      right: { x: box.x + box.w - 16, y: box.y, w: 16, h: 16 },
      thumb: { x: box.x + 16 + at, y: box.y, w: thumbW, h: 16 },
    }
  }
  return {
    header,
    clear: header ? { x: w - 30, y: 4, w: 22, h: 20 } : null,
    divider: header ? 27 : null,
    info,
    level: show.timeLevel && info ? { x: w - pad - 84, y: info.y, w: 84, h: 24 } : null,
    groupY,
    labelY,
    bars,
    periods,
    first: start,
    scrollbar,
  }
}

export type TimelineLook = {
  frame: string
  background: string
  heading: string
  divider: string
  /** The selection label. */
  label: string
  /** Year and period labels, and the time level. */
  muted: string
  /** Picked bars: a gradient top to bottom, and the lines between them. */
  picked: { top: string; bottom: string; line: string }
  unpicked: { fill: string; line: string }
  scroll: { track: string; thumb: string; frame: string }
}

const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const

/** A timeline style's colours under `theme` (an unknown style looks like Light 1). */
export function timelineLook(style: string, theme: Theme): TimelineLook {
  const m = /^TimeSlicerStyle(Light|Dark)(\d)$/.exec(style)
  const dark = m?.[1] === "Dark"
  const n = Math.min(6, Math.max(1, Number(m?.[2] ?? 1)))
  const accent = (theme.colors[ACCENTS[n - 1]] ?? "156082").toUpperCase()
  const hex = (tint: number) => "#" + (tint ? applyTint(accent, tint) : accent)
  return {
    frame: hex(-0.25),
    background: "#FFFFFF",
    heading: "#000000",
    divider: hex(-0.25),
    label: hex(-0.25),
    muted: "#7F7F7F",
    picked: dark
      ? { top: hex(-0.25), bottom: hex(-0.6), line: hex(0.4) }
      : { top: hex(0.6), bottom: hex(-0.25), line: hex(0.4) },
    unpicked: { fill: "#D9D9D9", line: "#FFFFFF" },
    scroll: { track: hex(0.8), thumb: "#FFFFFF", frame: "#A6A6A6" },
  }
}
