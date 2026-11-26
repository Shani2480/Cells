import { headerFooterEdges, headerFooterKind, headerFooterOf, headerPictureId, type PrintPage } from "@cells/core"
import { pagePlace } from "@/sheet/header-footer"
import type { SheetController } from "@/sheet/controller"
import type { PageAxis } from "@/sheet/page-map"
import { drawHeaderSection, sectionPicture, type SectionPainter } from "@/print/header-pictures"
import type { GridLayout } from "./layout"
import { CELL_FONT_FAMILY, type GridTheme } from "./theme"
import type { Painter } from "./draw"

/**
 * View › Page Layout, as Excel draws it: the sheet as sheets of paper on a grey desk. The cells sit inside each
 * page's margins; the header and footer show in the top and bottom margins. Cells are drawn first, then this covers
 * what falls between the pages and in the margins.
 */

/** One page on screen: its paper and its printable part, in canvas pixels. */
export type PageRect = {
  rowBand: number
  colBand: number
  x: number
  y: number
  w: number
  h: number
  body: { x: number; y: number; w: number; h: number }
}

/** The pages in view. */
export function visiblePages(layout: GridLayout): PageRect[] {
  const pages = layout.pages
  if (!pages) return []
  const { clip } = layout
  const along = (axis: PageAxis, origin: number, scroll: number, lo: number, hi: number) =>
    axis.bands.flatMap((band, i) => {
      const start = origin + band.start - scroll
      return start + axis.page < lo || start > hi ? [] : [{ i, start, band }]
    })
  const rows = along(pages.rows, layout.originY, layout.scrollY, clip.y0, clip.y1)
  const cols = along(pages.cols, layout.originX, layout.scrollX, clip.x0, clip.x1)
  const out: PageRect[] = []
  for (const r of rows)
    for (const c of cols)
      out.push({
        rowBand: r.i,
        colBand: c.i,
        x: c.start,
        y: r.start,
        w: pages.cols.page,
        h: pages.rows.page,
        body: {
          x: c.start + pages.cols.margin,
          y: r.start + pages.rows.margin,
          w: pages.cols.body,
          h: pages.rows.body,
        },
      })
  return out
}

/** A page's number as printing numbers it (down, then over, unless the setup says over, then down). */
function pageNumber(controller: SheetController, place: number) {
  const p = controller.print.current()
  return p.pages[place]?.number ?? (p.pages[0]?.number ?? 1) + place
}

/** The desk, the margins, each page's edge, and the header and footer (after the cells). */
export function drawPageLayout(p: Painter, controller: SheetController, layout: GridLayout, theme: GridTheme) {
  const map = layout.pages
  if (!map) return
  const { clip } = layout
  const pages = visiblePages(layout)
  p.save()
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  p.clip()

  // The desk: everything but the pages.
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  for (const page of pages) p.rect(page.x, page.y, page.w, page.h)
  p.fillStyle = theme.pageDesk
  p.fill("evenodd")

  // The margins: the paper less its printable part (cells and their text stop at the margin, as on paper).
  p.beginPath()
  for (const page of pages) {
    p.rect(page.x, page.y, page.w, page.h)
    p.rect(page.body.x, page.body.y, page.body.w, page.body.h)
  }
  p.fillStyle = theme.background
  p.fill("evenodd")

  p.strokeStyle = theme.pageEdge
  p.lineWidth = 1
  for (const page of pages) p.strokeRect(page.x - 0.5, page.y - 0.5, page.w + 1, page.h + 1)

  // Header and footer: 11 pt, scaled with the sheet as it prints.
  const setup = controller.sheet.pageSetup
  const pagination = controller.print.current()
  const scale = pagination.scale || 1
  const pxPerIn = 96 / scale
  // 11 pt at the print scale, on a page drawn 1/scale as big: the same size on screen at any scale (without Scale
  // with Document, 11 pt on the paper).
  const size = 11 / 0.75 / (setup.scaleWithDoc ? 1 : scale)
  p.font = `${size}px ${CELL_FONT_FAMILY}`
  p.textBaseline = "alphabetic"
  const ctx = { ...controller.printFile, now: new Date(), total: Math.max(1, pagination.pages.length) }
  const editing = controller.headerFooter.editing
  for (const page of pages) {
    const place = pagePlace(controller, page.rowBand, page.colBand)
    const number = pageNumber(controller, place)
    const printPage: PrintPage = { area: 0, rows: [0, 0], cols: [0, 0], titleRows: null, titleCols: null, number }
    const edges = headerFooterEdges(setup, page.w / pxPerIn)
    const left = page.x + edges.left * pxPerIn
    const right = page.x + edges.right * pxPerIn
    for (const where of ["header", "footer"] as const) {
      // The one being edited shows in its boxes instead.
      if (editing?.where === where && editing.rowBand === page.rowBand && editing.colBand === page.colBand) continue
      const hf = headerFooterOf(setup, where, place)
      const empty = !hf.left && !hf.center && !hf.right
      if (empty) {
        // Excel's hint on the first page's empty header or footer (and the second's, when that one differs), faint,
        // where a click adds one.
        if (place > (setup.differentFirst || setup.differentOddEven ? 1 : 0)) continue
        p.fillStyle = theme.pageHint
        p.textAlign = "center"
        const y =
          where === "header"
            ? page.y + setup.margins.header * pxPerIn + size
            : page.y + page.h - setup.margins.footer * pxPerIn
        p.fillText(where === "header" ? "Add header" : "Add footer", (left + right) / 2, y)
        continue
      }
      p.fillStyle = theme.text
      const kind = headerFooterKind(setup, place)
      for (const [section, align, x] of [
        ["left", "left", left],
        ["center", "center", (left + right) / 2],
        ["right", "right", right],
      ] as const) {
        const text = hf[section]
        if (!text) continue
        drawHeaderSection(p as SectionPainter, {
          text,
          picture: sectionPicture(setup, headerPictureId(where, kind, section), text),
          align,
          x,
          y:
            where === "header"
              ? page.y + setup.margins.header * pxPerIn
              : page.y + page.h - setup.margins.footer * pxPerIn,
          where,
          size,
          // Points on the page as drawn (1/scale as big; without Scale with Document, a point on paper).
          pt: size / 11,
          page: printPage,
          sheet: controller.sheet,
          ctx,
          // Pictures load once; the view draws again when one has.
          onImage: () => controller.changed(),
        })
      }
    }
  }
  p.textAlign = "left"
  p.restore()
}

/** The page a point is on, if any. */
export function pageAt(layout: GridLayout, x: number, y: number): PageRect | null {
  return visiblePages(layout).find((p) => x >= p.x && x <= p.x + p.w && y >= p.y && y <= p.y + p.h) ?? null
}

/** Which of a header's three sections a point is over (thirds of the printable width). */
export function sectionAt(page: PageRect, x: number): "left" | "center" | "right" {
  const t = (x - page.body.x) / page.body.w
  return t < 1 / 3 ? "left" : t < 2 / 3 ? "center" : "right"
}

/** What part of a page a point is on: its header or footer margin, another margin or the desk, or its cells. */
export function pagePartAt(layout: GridLayout, x: number, y: number) {
  if (!layout.pages) return null
  for (const page of visiblePages(layout)) {
    if (x < page.x || x > page.x + page.w || y < page.y || y > page.y + page.h) continue
    const b = page.body
    if (x >= b.x && x <= b.x + b.w && y >= b.y && y <= b.y + b.h) return null
    if (y < b.y) return "header" as const
    if (y > b.y + b.h) return "footer" as const
    return "margin" as const
  }
  return "desk" as const
}
