import {
  slicerLayout,
  slicerLook,
  type SheetSlicer,
  type SlicerButtonLook,
  type SlicerItem,
  type Theme,
} from "@cells/core"

/** What a slicer shows: its buttons, whether its column is filtered, and whether Multi-Select is on. */
export type SlicerView = { items: SlicerItem[]; filtered: boolean; multi: boolean }

/**
 * A slicer drawn as Excel 365 draws it (sampled from its pictures): a frame in the style's colour, the caption in
 * bold with the Multi-Select and Clear Filter buttons, a line under the header, and rounded buttons coloured by
 * whether they're picked and whether the table's other filters leave rows for them.
 */

type Measure = (text: string, font: string) => number

let source: ((o: SheetSlicer) => SlicerView) | null = null
/** Where slicers get their buttons (the controller, which knows the tables). */
export function setSlicerSource(fn: ((o: SheetSlicer) => SlicerView) | null) {
  source = fn
}
export const slicerView = (o: SheetSlicer): SlicerView => source?.(o) ?? { items: [], filtered: false, multi: false }

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
const n = (v: number) => Math.round(v * 100) / 100

/** `text` cut to `width` with an ellipsis, as Excel cuts long captions and items. */
export function fit(text: string, width: number, font: string, measure: Measure): string {
  if (measure(text, font) <= width) return text
  let lo = 0
  let hi = text.length
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2)
    if (measure(text.slice(0, mid) + "…", font) <= width) lo = mid
    else hi = mid - 1
  }
  return lo ? text.slice(0, lo) + "…" : ""
}

export function slicerSvg(
  o: SheetSlicer,
  view: SlicerView,
  w: number,
  h: number,
  theme: Theme,
  measure: Measure,
): string {
  const look = slicerLook(o.style, theme)
  const layout = slicerLayout(o, w, h, view.items.length)
  const family = `"${theme.fonts.minor}", "Aptos", "Calibri", sans-serif`
  const heading = `700 14.67px ${family}`
  const body = `14.67px ${family}`
  const out: string[] = [
    `<rect x="0.5" y="0.5" width="${n(w - 1)}" height="${n(h - 1)}" fill="${look.background}" stroke="${look.frame}"/>`,
  ]
  if (layout.header) {
    const room = (layout.multi?.x ?? w) - 9 - 6
    out.push(
      `<text x="9" y="20" font-family='${family}' font-weight="700" font-size="14.67" fill="${look.heading}">${esc(fit(o.caption, room, heading, measure))}</text>`,
    )
    if (layout.multi) out.push(multiIcon(layout.multi, view.multi))
    if (layout.clear) out.push(clearIcon(layout.clear, view.filtered))
    if (look.divider && layout.divider !== null)
      out.push(
        `<line x1="9" y1="${layout.divider + 0.5}" x2="${n(w - 9)}" y2="${layout.divider + 0.5}" stroke="${look.divider}"/>`,
      )
  }
  layout.buttons.forEach((b, i) => {
    const item = view.items[layout.first + i]
    if (!item) return
    const s: SlicerButtonLook = item.selected
      ? item.data
        ? look.selected
        : look.selectedNoData
      : item.data
        ? look.unselected
        : look.unselectedNoData
    out.push(
      `<rect x="${n(b.x + 0.5)}" y="${n(b.y + 0.5)}" width="${n(b.w - 1)}" height="${n(b.h - 1)}" rx="2.5" fill="${s.fill}" stroke="${s.border}"/>`,
      `<text x="${n(b.x + 9.5)}" y="${n(b.y + b.h / 2 + 4.5)}" font-family='${family}' font-size="14.67" fill="${s.text}">${esc(fit(item.text, b.w - 16, body, measure))}</text>`,
    )
  })
  if (layout.scroll && layout.buttons.length) {
    const top = layout.buttons[0].y
    const bottom = layout.buttons[layout.buttons.length - 1].y + layout.buttons[layout.buttons.length - 1].h
    const track = bottom - top
    const thumb = Math.max(12, (track * layout.scroll.shown) / layout.scroll.total)
    const columns = Math.max(1, o.columns)
    const at = Math.floor(layout.first / columns) / Math.max(1, layout.scroll.total - layout.scroll.shown)
    out.push(
      `<rect x="${n(w - 15)}" y="${n(top)}" width="7" height="${n(track)}" rx="3.5" fill="#F0F0F0"/>`,
      `<rect x="${n(w - 15)}" y="${n(top + (track - thumb) * at)}" width="7" height="${n(thumb)}" rx="3.5" fill="#A6A6A6"/>`,
    )
  }
  return out.join("")
}

/** Multi-Select: three ticks beside three lines; a light square behind it while it's on. */
function multiIcon(b: { x: number; y: number; w: number; h: number }, on: boolean): string {
  const x = b.x
  const y = b.y
  const ticks = [3, 8, 13]
    .map(
      (dy) =>
        `<polyline points="${x + 1.5},${y + dy} ${x + 3.5},${y + dy + 2} ${x + 7},${y + dy - 2}" fill="none" stroke="#505050" stroke-width="1.1"/>` +
        `<line x1="${x + 9}" y1="${y + dy + 0.5}" x2="${x + 17}" y2="${y + dy + 0.5}" stroke="#505050" stroke-width="1.1"/>`,
    )
    .join("")
  const back = on
    ? `<rect x="${x - 2.5}" y="${y - 2.5}" width="${b.w + 4}" height="${b.h + 3}" rx="2" fill="#CDE6F7" stroke="#92C0E0"/>`
    : ""
  return back + ticks
}

/** Clear Filter: a funnel, with a red cross while the column is filtered (greyed out otherwise). */
export function clearIcon(b: { x: number; y: number; w: number; h: number }, active: boolean): string {
  const x = b.x
  const y = b.y
  const ink = active ? "#505050" : "#B4B4B4"
  const funnel = `<path d="M${x + 1} ${y + 1.5} H${x + 15} L${x + 9.5} ${y + 8} V${y + 15} L${x + 6.5} ${y + 13} V${y + 8} Z" fill="none" stroke="${ink}" stroke-width="1.1" stroke-linejoin="round"/>`
  const cross = `<path d="M${x + 12} ${y + 10} L${x + 18} ${y + 16} M${x + 18} ${y + 10} L${x + 12} ${y + 16}" stroke="${active ? "#E81123" : "#D0D0D0"}" stroke-width="1.4"/>`
  return funnel + cross
}
