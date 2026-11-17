import {
  sparkColorHex,
  sparklineBounds,
  sparklinePoints,
  type Sheet,
  type SparklineGroup,
  type Theme,
} from "@cells/core"
import type { CellPainter } from "./draw-cells"

/**
 * Sparklines drawn in their cells. What each one plots is worked out once per workbook version (a group's shared
 * axis needs every sparkline of the group), then each frame draws the ones in view.
 */

export type SparkItem = {
  group: SparklineGroup
  values: (number | null)[]
  x: number[]
  bounds: { min: number; max: number } | null
}

type Cache = { sheet: Sheet; version: number; items: Map<number, SparkItem> }
let cache: Cache | null = null

const key = (row: number, col: number) => row * 16384 + col

/** The sparklines of a sheet by cell (row * 16384 + col). */
export function sparkItems(sheets: readonly Sheet[], sheet: Sheet, version: number): Map<number, SparkItem> {
  if (cache && cache.sheet === sheet && cache.version === version) return cache.items
  const items = new Map<number, SparkItem>()
  for (const group of sheet.sparklines) {
    const points = group.sparklines.map((s) => sparklinePoints(sheets, group, s))
    const bounds = sparklineBounds(group, points)
    group.sparklines.forEach((s, i) => items.set(key(s.row, s.col), { group, ...points[i], bounds: bounds[i] }))
  }
  cache = { sheet, version, items }
  return items
}

export const sparkAt = (items: Map<number, SparkItem>, row: number, col: number) => items.get(key(row, col))

type Rect = { x: number; y: number; w: number; h: number }

/** Excel keeps a margin inside the cell: about 4 px across and a sixth of the height above and below. */
const MARGIN_X = 4
const marginY = (h: number) => Math.max(2, Math.round(h * 0.17))
/** Columns and win/loss bars take three quarters of their slot. */
const BAR = 0.75

/** One sparkline in `rect` (sheet pixels at 100%). `paint` adjusts colours for dark mode. */
export function drawSparkline(p: CellPainter, rect: Rect, item: SparkItem, theme: Theme, paint: (c: string) => string) {
  const { group: g, values, x } = item
  const n = values.length
  if (!n || !item.bounds) return
  const color = (c: SparklineGroup["colors"][keyof SparklineGroup["colors"]]) => paint(sparkColorHex(c, theme))
  const left = rect.x + MARGIN_X
  const width = Math.max(1, rect.w - MARGIN_X * 2)
  const top = rect.y + marginY(rect.h)
  const height = Math.max(1, rect.h - marginY(rect.h) * 2)
  p.save()
  p.beginPath()
  p.rect(rect.x, rect.y, rect.w, rect.h)
  p.clip()

  const numbers = values.filter((v): v is number => v !== null)
  const high = numbers.length ? Math.max(...numbers) : 0
  const low = numbers.length ? Math.min(...numbers) : 0
  const firstAt = values.findIndex((v) => v !== null)
  let lastAt = -1
  for (let i = n - 1; i >= 0; i--)
    if (values[i] !== null) {
      lastAt = i
      break
    }
  /** The special colour a point takes, if any (high and low over first and last, over negative). */
  const pointColor = (i: number): string | null => {
    const v = values[i]
    if (v === null) return null
    if (g.high && v === high) return color(g.colors.high)
    if (g.low && v === low) return color(g.colors.low)
    if (g.first && i === firstAt) return color(g.colors.first)
    if (g.last && i === lastAt) return color(g.colors.last)
    if (g.negative && v < 0) return color(g.colors.negative)
    return null
  }

  if (g.type === "stacked") {
    // Win/Loss: up for a win, down for a loss, nothing for zero; the axis in the middle.
    const slot = width / n
    const bar = Math.max(1, slot * BAR)
    const mid = top + height / 2
    for (let i = 0; i < n; i++) {
      const v = values[i]
      if (v === null || v === 0) continue
      p.fillStyle = pointColor(i) ?? (v < 0 && g.negative ? color(g.colors.negative) : color(g.colors.series))
      const cx = left + (g.rightToLeft ? n - 1 - i : i) * slot + (slot - bar) / 2
      if (v > 0) p.fillRect(cx, top, bar, mid - top - 0.5)
      else p.fillRect(cx, mid + 0.5, bar, top + height - mid - 0.5)
    }
    if (g.displayXAxis) axisLine(p, left, width, mid, color(g.colors.axis))
    p.restore()
    return
  }

  let { min, max } = item.bounds
  if (g.type === "column") {
    // Columns grow from zero when the range spans it; otherwise from just below the smallest value.
    if (max === min) {
      min = Math.min(0, min)
      max = Math.max(0, max)
    }
  }
  const span = max - min || 1
  const yOf = (v: number) => top + height - ((Math.min(max, Math.max(min, v)) - min) / span) * height
  const zero = min <= 0 && max >= 0 ? yOf(0) : null

  if (g.type === "column") {
    const slot = width / n
    const bar = Math.max(1, slot * BAR)
    const base = zero ?? (min > 0 ? top + height : top)
    for (let i = 0; i < n; i++) {
      const v = values[i]
      if (v === null) continue
      p.fillStyle = pointColor(i) ?? color(g.colors.series)
      const cx = left + x[i] * (width - slot) + (slot - bar) / 2
      let y = yOf(v)
      // The smallest value still shows as a sliver, as in Excel.
      if (Math.abs(y - base) < 1) y = v < 0 ? base + 1 : base - 1
      p.fillRect(cx, Math.min(y, base), bar, Math.max(1, Math.abs(base - y)))
    }
    if (g.displayXAxis && zero !== null) axisLine(p, left, width, zero, color(g.colors.axis))
    p.restore()
    return
  }

  // Line: joined points, broken at gaps (or joined over them with "span").
  if (g.displayXAxis && zero !== null) axisLine(p, left, width, zero, color(g.colors.axis))
  const lineWidth = Math.max(0.5, (g.lineWeight * 96) / 72)
  p.strokeStyle = color(g.colors.series)
  p.lineWidth = lineWidth
  p.beginPath()
  let open = false
  const px = (i: number) => left + x[i] * width
  for (let i = 0; i < n; i++) {
    const v = values[i]
    if (v === null) {
      if (g.emptyAs !== "span") open = false
      continue
    }
    if (open) p.lineTo(px(i), yOf(v))
    else p.moveTo(px(i), yOf(v))
    open = true
  }
  if (numbers.length === 1) {
    const i = firstAt
    p.moveTo(px(i) - 0.5, yOf(values[i]!))
    p.lineTo(px(i) + 0.5, yOf(values[i]!))
  }
  p.stroke()
  // Markers and special points are round dots, a little wider than the line.
  const radius = Math.max(1.5, (lineWidth + 1.5) / 2)
  for (let i = 0; i < n; i++) {
    const v = values[i]
    if (v === null) continue
    const special = pointColor(i)
    if (!special && !g.markers) continue
    p.fillStyle = special ?? color(g.colors.markers)
    dot(p, px(i), yOf(v), radius)
  }
  p.restore()
}

function dot(p: CellPainter, cx: number, cy: number, r: number) {
  p.beginPath()
  const steps = 16
  for (let i = 0; i <= steps; i++) {
    const a = (i / steps) * Math.PI * 2
    if (i === 0) p.moveTo(cx + r, cy)
    else p.lineTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r)
  }
  p.closePath()
  p.fill()
}

function axisLine(p: CellPainter, left: number, width: number, y: number, color: string) {
  p.strokeStyle = color
  p.lineWidth = 1
  p.beginPath()
  p.moveTo(left, Math.round(y) + 0.5)
  p.lineTo(left + width, Math.round(y) + 0.5)
  p.stroke()
}
