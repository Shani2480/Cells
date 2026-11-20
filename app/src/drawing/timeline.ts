import {
  timelineLayout,
  timelineLook,
  type SheetTimeline,
  type Theme,
  type TimelineLevel,
  type TimelinePeriod,
} from "@cells/core"
import { clearIcon, fit } from "./slicer"

/**
 * A timeline drawn as Excel 365 draws it (sampled from its pictures of Light 1 and Dark 1): the caption over a line,
 * the selection label and the time level, the year (or month) above the period labels, a bar per period — picked
 * ones in the style's colour — and the scroll bar.
 */

type Measure = (text: string, font: string) => number

/** What a timeline shows: its periods at its level, the dates picked, and the first period shown. */
export type TimelineView = {
  periods: TimelinePeriod[]
  picked: { from: number; to: number } | null
  first: number
  label: string
}

let source: ((o: SheetTimeline) => TimelineView) | null = null
/** Where timelines get their periods (the controller, which knows the PivotTables). */
export function setTimelineSource(fn: ((o: SheetTimeline) => TimelineView) | null) {
  source = fn
}
export const timelineView = (o: SheetTimeline): TimelineView =>
  source?.(o) ?? { periods: [], picked: null, first: 0, label: "All Periods" }

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
const n = (v: number) => Math.round(v * 100) / 100
export const LEVEL_NAMES: Record<TimelineLevel, string> = {
  years: "YEARS",
  quarters: "QUARTERS",
  months: "MONTHS",
  days: "DAYS",
}

export function timelineSvg(
  o: SheetTimeline,
  view: TimelineView,
  w: number,
  h: number,
  theme: Theme,
  measure: Measure,
): string {
  const look = timelineLook(o.style, theme)
  const layout = timelineLayout(o, o.level, w, view.periods.length, view.first)
  const family = `"${theme.fonts.minor}", "Aptos", "Calibri", sans-serif`
  const heading = `700 14.67px ${family}`
  const body = `14.67px ${family}`
  const small = `12px ${family}`
  const id = `tl${o.id}`
  const out: string[] = [
    `<defs><linearGradient id="${id}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${look.picked.top}"/><stop offset="1" stop-color="${look.picked.bottom}"/></linearGradient>` +
      `<clipPath id="${id}c"><rect x="${layout.bars.x}" y="0" width="${n(layout.bars.w)}" height="${n(h)}"/></clipPath></defs>`,
    `<rect x="0.5" y="0.5" width="${n(w - 1)}" height="${n(h - 1)}" fill="${look.background}" stroke="${look.frame}"/>`,
  ]
  if (layout.header) {
    out.push(
      `<text x="9" y="20" font-family='${family}' font-weight="700" font-size="14.67" fill="${look.heading}">${esc(fit(o.caption, w - 48, heading, measure))}</text>`,
    )
    if (layout.clear) out.push(clearIcon(layout.clear, view.picked !== null))
    out.push(
      `<line x1="8" y1="${layout.divider! + 0.5}" x2="${n(w - 8)}" y2="${layout.divider! + 0.5}" stroke="${look.divider}"/>`,
    )
  }
  if (layout.info) {
    const y = layout.info.y + 17
    if (o.selectionLabel)
      out.push(
        `<text x="${layout.info.x}" y="${y}" font-family='${family}' font-size="14.67" fill="${look.label}">${esc(fit(view.label, layout.info.w - (layout.level ? 90 : 0), body, measure))}</text>`,
      )
    if (layout.level) {
      const text = LEVEL_NAMES[o.level]
      const right = layout.level.x + layout.level.w
      out.push(
        `<text x="${right - 16}" y="${y}" text-anchor="end" font-family='${family}' font-size="12" fill="${look.muted}">${text}</text>`,
        `<path d="M${right - 11} ${y - 7} h8 l-4 4 z" fill="${look.muted}"/>`,
      )
    }
  }
  const bars: string[] = []
  const labels: string[] = []
  layout.periods.forEach((b, i) => {
    const p = view.periods[layout.first + i]
    if (!p) return
    const on = view.picked !== null && p.start <= view.picked.to && p.end > view.picked.from
    bars.push(
      `<rect x="${n(b.x)}" y="${n(b.y)}" width="${n(b.w)}" height="${n(b.h)}" fill="${on || !view.picked ? `url(#${id})` : look.unpicked.fill}"/>`,
      `<line x1="${n(b.x + b.w - 0.5)}" y1="${n(b.y)}" x2="${n(b.x + b.w - 0.5)}" y2="${n(b.y + b.h)}" stroke="${on || !view.picked ? look.picked.line : look.unpicked.line}"/>`,
    )
    labels.push(
      `<text x="${n(b.x + 3)}" y="${n(layout.labelY + 4)}" font-family='${family}' font-size="12" fill="${look.muted}">${esc(fit(p.label, b.w - 4, small, measure))}</text>`,
    )
    const group = i === 0 ? (p.group ?? groupOf(view.periods, layout.first)) : p.group
    if (group)
      labels.push(
        `<text x="${n(b.x + 3)}" y="${n(layout.groupY + 4)}" font-family='${family}' font-size="12" fill="${look.muted}">${esc(group)}</text>`,
      )
  })
  out.push(`<g clip-path="url(#${id}c)">${bars.join("")}${labels.join("")}</g>`)
  const sb = layout.scrollbar
  if (sb) {
    const arrow = (b: { x: number; y: number }, left: boolean) =>
      `<rect x="${n(b.x + 0.5)}" y="${n(b.y + 0.5)}" width="15" height="15" fill="#FFFFFF" stroke="${look.scroll.frame}"/>` +
      `<path d="${left ? `M${b.x + 10} ${b.y + 4} l-5 4 l5 4 z` : `M${b.x + 6} ${b.y + 4} l5 4 l-5 4 z`}" fill="#505050"/>`
    out.push(
      `<rect x="${n(sb.box.x + 16)}" y="${n(sb.box.y + 4)}" width="${n(sb.box.w - 32)}" height="8" fill="${look.scroll.track}"/>`,
      `<rect x="${n(sb.thumb.x + 0.5)}" y="${n(sb.thumb.y + 0.5)}" width="${n(sb.thumb.w - 1)}" height="15" rx="2" fill="${look.scroll.thumb}" stroke="${look.scroll.frame}"/>`,
      arrow(sb.left, true),
      arrow(sb.right, false),
    )
  }
  return out.join("")
}

/** The group label (year, or month for days) of the period at `i`, from the last period that names one. */
function groupOf(periods: readonly TimelinePeriod[], i: number): string | null {
  for (let k = i; k >= 0; k--) if (periods[k].group) return periods[k].group
  return null
}
