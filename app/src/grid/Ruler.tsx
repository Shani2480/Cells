import type { ReactNode } from "react"
import type { SheetController } from "@/sheet/controller"
import type { GridLayout } from "./layout"

/** The ruler's thickness (screen pixels; it doesn't zoom). */
export const RULER = 18

/** View › Ruler: shown above and left of the sheet in Page Layout only. */
export function rulerShown(controller: SheetController): boolean {
  return controller.view === "pageLayout" && controller.showRuler
}

/** Inches where the US measures in them, centimetres elsewhere (Excel follows the system's units). */
export function rulerUnits(locale = typeof navigator === "undefined" ? "en-US" : navigator.language) {
  return /^en-(US|LR)$|^my\b/i.test(locale) ? "in" : "cm"
}

/**
 * Page Layout's ruler along the columns (above the headings) or the rows (left of them): for each page in view its
 * margins shaded and its printable part white, with a number every inch or centimetre counted from the margin.
 */
export function Ruler({
  controller,
  layout,
  zoom,
  axis,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  axis: "rows" | "cols"
}) {
  const map = layout.pages
  if (!map || !rulerShown(controller)) return null
  const rows = axis === "rows"
  const pages = rows ? map.rows : map.cols
  const origin = rows ? layout.originY : layout.originX
  const scroll = rows ? layout.scrollY : layout.scrollX
  const head = (rows ? layout.headerHeight : layout.rowHeaderWidth) * zoom
  const extent = (rows ? layout.height : layout.width) * zoom
  const scale = controller.print.current().scale || 1
  const units = rulerUnits()
  // Sheet pixels per unit on the page (a page prints at `scale`, so it's drawn 1/scale as big).
  const unit = (units === "in" ? 96 : 96 / 2.54) / scale
  const steps = units === "in" ? 4 : 2
  const along = (at: number) => (origin + at - scroll) * zoom - head
  const box = (from: number, size: number) => (rows ? { top: from, height: size } : { left: from, width: size })

  const parts: ReactNode[] = []
  pages.bands.forEach((band, i) => {
    const start = along(band.start)
    const size = pages.page * zoom
    if (start + size < 0 || start > extent - head) return
    const body = start + pages.margin * zoom
    parts.push(
      <div key={`p${i}`} className="ruler-page" style={box(start, size)} />,
      <div key={`b${i}`} className="ruler-body" style={box(body, pages.body * zoom)} />,
    )
    // Ticks from the left (or top) margin both ways, across the whole page.
    const first = -Math.floor((pages.margin / unit) * steps)
    const last = Math.floor(((pages.page - pages.margin) / unit) * steps)
    for (let k = first; k <= last; k++) {
      const at = body + (k / steps) * unit * zoom
      if (k % steps === 0 && k !== 0)
        parts.push(
          <span key={`l${i}:${k}`} className="ruler-label" style={rows ? { top: at } : { left: at }}>
            {Math.abs(k / steps)}
          </span>,
        )
      else
        parts.push(
          <span
            key={`t${i}:${k}`}
            className={`ruler-tick${k % 2 ? "" : " is-half"}`}
            style={rows ? { top: at } : { left: at }}
          />,
        )
    }
  })
  return (
    <div className={`ruler ${rows ? "is-rows" : "is-cols"}`} aria-hidden="true">
      <div className="ruler-track" style={rows ? { top: head } : { left: head }}>
        {parts}
      </div>
    </div>
  )
}
