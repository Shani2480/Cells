import type { Range } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { colX, rowY, type GridLayout } from "./layout"
import type { GridTheme } from "./theme"
import type { Painter } from "./draw"

/**
 * View › Page Break Preview, as Excel draws it: what doesn't print is shaded grey, each page is outlined in blue
 * (the print area's edge thick, automatic breaks dashed, breaks set by hand solid) and has "Page N" across it.
 * Dragging a blue line moves the break, or the print area's edge.
 */

/** A blue line of the preview: a page break or an edge of a print area, before row or column `index`. */
export type BreakLine = {
  axis: "rows" | "cols"
  index: number
  /** "edge": the print area's border (dragging it changes the print area); "break": between pages. */
  kind: "edge" | "break"
  manual: boolean
  /** Which print area, and the span the line runs along (rows for a column line, columns for a row line). */
  area: number
  from: number
  to: number
}

/** Every blue line of the active sheet's pages. */
export function breakLines(controller: SheetController): { lines: BreakLine[]; pages: PageBox[] } {
  const p = controller.print.current()
  const setup = controller.sheet.pageSetup
  const lines: BreakLine[] = []
  const seen = new Set<string>()
  const add = (line: BreakLine) => {
    const key = `${line.area}:${line.axis}:${line.index}`
    if (seen.has(key)) return
    seen.add(key)
    lines.push(line)
  }
  p.areas.forEach((a, area) => {
    add({ axis: "rows", index: a.r0, kind: "edge", manual: true, area, from: a.c0, to: a.c1 })
    add({ axis: "rows", index: a.r1 + 1, kind: "edge", manual: true, area, from: a.c0, to: a.c1 })
    add({ axis: "cols", index: a.c0, kind: "edge", manual: true, area, from: a.r0, to: a.r1 })
    add({ axis: "cols", index: a.c1 + 1, kind: "edge", manual: true, area, from: a.r0, to: a.r1 })
  })
  for (const page of p.pages) {
    const a = p.areas[page.area]
    if (page.rows[0] > a.r0)
      add({
        axis: "rows",
        index: page.rows[0],
        kind: "break",
        manual: setup.rowBreaks.includes(page.rows[0]),
        area: page.area,
        from: a.c0,
        to: a.c1,
      })
    if (page.cols[0] > a.c0)
      add({
        axis: "cols",
        index: page.cols[0],
        kind: "break",
        manual: setup.colBreaks.includes(page.cols[0]),
        area: page.area,
        from: a.r0,
        to: a.r1,
      })
  }
  const pages = p.pages.map((page) => ({
    range: { r0: page.rows[0], r1: page.rows[1], c0: page.cols[0], c1: page.cols[1] },
    number: page.number,
  }))
  return { lines, pages }
}

/** The "Page N" labels' size in sheet pixels at 100%. */
const PAGE_LABEL_SIZE = 40

export type PageBox = { range: Range; number: number }

/** The preview over one pane: shading, page numbers and the blue lines (cells are already drawn). */
export function drawBreakPreview(p: Painter, controller: SheetController, layout: GridLayout, theme: GridTheme) {
  if (controller.view !== "pageBreakPreview" || !layout.rows.length || !layout.cols.length) return
  const sheet = controller.sheet
  const { lines, pages } = breakLines(controller)
  const { clip } = layout
  const rect = (r: Range) => {
    const x = colX(sheet, layout, r.c0)
    const y = rowY(sheet, layout, r.r0)
    return {
      x,
      y,
      w: colX(sheet, layout, r.c1) + sheet.cols.size(r.c1) - x,
      h: rowY(sheet, layout, r.r1) + sheet.rows.size(r.r1) - y,
    }
  }
  p.save()
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  p.clip()

  // Grey over what doesn't print: the pane, less each page.
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  for (const page of pages) {
    const b = rect(page.range)
    p.rect(b.x, b.y, b.w, b.h)
  }
  p.fillStyle = theme.breakPreviewShade
  p.fill("evenodd")

  // "Page N" across each page, big and faint.
  p.textAlign = "center"
  p.textBaseline = "middle"
  p.fillStyle = theme.breakPreviewLabel
  for (const page of pages) {
    const b = rect(page.range)
    const cx = b.x + b.w / 2
    const cy = b.y + b.h / 2
    if (cx < clip.x0 - b.w || cx > clip.x1 + b.w || cy < clip.y0 - b.h || cy > clip.y1 + b.h) continue
    // One size for every page (it scales with the zoom, as the cells do), smaller only on a tiny page.
    const size = Math.max(10, Math.min(PAGE_LABEL_SIZE, b.w / 5, b.h / 2))
    p.font = `600 ${Math.round(size)}px Aptos, "Segoe UI", system-ui, sans-serif`
    p.fillText(`Page ${page.number}`, cx, cy)
  }
  p.textBaseline = "alphabetic"
  p.textAlign = "left"

  // The blue lines.
  p.strokeStyle = theme.breakPreviewLine
  for (const line of lines) {
    p.lineWidth = line.kind === "edge" ? 3 : 2
    p.setLineDash(line.manual ? [] : [8, 4])
    p.beginPath()
    if (line.axis === "rows") {
      const y = Math.round(rowY(sheet, layout, line.index))
      const x0 = colX(sheet, layout, line.from)
      const x1 = colX(sheet, layout, line.to) + sheet.cols.size(line.to)
      p.moveTo(x0, y)
      p.lineTo(x1, y)
    } else {
      const x = Math.round(colX(sheet, layout, line.index))
      const y0 = rowY(sheet, layout, line.from)
      const y1 = rowY(sheet, layout, line.to) + sheet.rows.size(line.to)
      p.moveTo(x, y0)
      p.lineTo(x, y1)
    }
    p.stroke()
  }
  p.setLineDash([])
  p.restore()
}

/** How near a blue line a press grabs it (sheet pixels at 100%; the zoom is applied to it). */
const GRAB = 4

/** The blue line under a point of the canvas (sheet pixels), if any. */
export function breakLineAt(controller: SheetController, layout: GridLayout, x: number, y: number): BreakLine | null {
  if (controller.view !== "pageBreakPreview") return null
  if (x < layout.rowHeaderWidth || y < layout.headerHeight) return null
  const sheet = controller.sheet
  const reach = GRAB / controller.zoom
  let best: BreakLine | null = null
  let bestDistance = reach
  for (const line of breakLines(controller).lines) {
    if (line.axis === "rows") {
      const ly = rowY(sheet, layout, line.index)
      const x0 = colX(sheet, layout, line.from)
      const x1 = colX(sheet, layout, line.to) + sheet.cols.size(line.to)
      const d = Math.abs(y - ly)
      if (x >= x0 - reach && x <= x1 + reach && d <= bestDistance) [best, bestDistance] = [line, d]
    } else {
      const lx = colX(sheet, layout, line.index)
      const y0 = rowY(sheet, layout, line.from)
      const y1 = rowY(sheet, layout, line.to) + sheet.rows.size(line.to)
      const d = Math.abs(x - lx)
      if (y >= y0 - reach && y <= y1 + reach && d <= bestDistance) [best, bestDistance] = [line, d]
    }
  }
  return best
}

/** The row or column boundary nearest a point, for where a dragged line lands. */
export function boundaryAt(controller: SheetController, layout: GridLayout, axis: "rows" | "cols", at: number) {
  const sheet = controller.sheet
  const lines = axis === "rows" ? layout.rows : layout.cols
  let best = lines[0]?.index ?? 0
  let bestDistance = Infinity
  for (const l of lines) {
    for (const [edge, index] of [
      [l.start, l.index],
      [l.start + l.size, l.index + 1],
    ] as const) {
      const d = Math.abs(at - edge)
      if (d < bestDistance) [best, bestDistance] = [index, d]
    }
  }
  return Math.max(0, Math.min(best, axis === "rows" ? sheet.rows.count : sheet.cols.count))
}
