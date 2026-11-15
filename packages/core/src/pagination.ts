import type { Range } from "./address"
import type { Axis } from "./axis"
import { paperInches, type PageSetup } from "./page-setup"
import type { Sheet } from "./sheet"

/**
 * Pages: how a sheet prints with its page setup, worked out as Excel does. Everything is in points (1/72 inch); the
 * sheet's sizes are pixels at 96 dpi (¾ point each).
 */

export const PT_PER_PX = 0.75

/** One printed page: the rows and columns of its print area it shows, and the titles repeated above and beside. */
export type PrintPage = {
  /** Which print area (index into Pagination.areas). */
  area: number
  rows: [number, number]
  cols: [number, number]
  titleRows: [number, number] | null
  titleCols: [number, number] | null
  /** The page number printed for &P. */
  number: number
}

export type Pagination = {
  pages: PrintPage[]
  /** What prints: the print area, or the cells in use. */
  areas: Range[]
  /** The scale everything prints at (1 = 100%), after Fit to. */
  scale: number
  /** The paper and the part inside the margins, in points. */
  paper: { width: number; height: number }
  body: { x: number; y: number; width: number; height: number }
}

export type PaginateOptions = {
  /** Instead of the print area (File › Print › Print Selection). */
  ranges?: Range[]
  /** The first page's number when the setup says automatic (pages of the sheets before it, plus one). */
  firstNumber?: number
  /** Row and column headings when they print: the row number column's width and the header row's height (pt). */
  headings?: { width: number; height: number }
  /** Ignore the print area (File › Print › Ignore Print Area). */
  ignorePrintArea?: boolean
  /** The paper when the setup names none (see defaultPaperSize). */
  defaultPaper?: number
}

/** The cells a sheet prints when it has no print area: everything with content or a format, and its charts. */
export function printedRange(sheet: Sheet): Range | null {
  let r = sheet.cells.usedBounds()
  for (const chart of sheet.charts) {
    const { from, to } = chart.anchor
    const box = { r0: from.row, c0: from.col, r1: to.row, c1: to.col }
    r = r
      ? {
          r0: Math.min(r.r0, box.r0),
          c0: Math.min(r.c0, box.c0),
          r1: Math.max(r.r1, box.r1),
          c1: Math.max(r.c1, box.c1),
        }
      : box
  }
  for (const m of sheet.merges.all) {
    if (!r || m.r0 > r.r1 || m.r1 < r.r0 || m.c0 > r.c1 || m.c1 < r.c0) continue
    r = { r0: Math.min(r.r0, m.r0), c0: Math.min(r.c0, m.c0), r1: Math.max(r.r1, m.r1), c1: Math.max(r.c1, m.c1) }
  }
  return r ? { ...r } : null
}

/** A line's size in points (hidden lines take none). */
const size = (axis: Axis, i: number) => (axis.isHidden(i) ? 0 : axis.size(i) * PT_PER_PX)
const span = (axis: Axis, a: number, b: number) => {
  let total = 0
  for (let i = a; i <= b; i++) total += size(axis, i)
  return total
}

/**
 * Splits lines a..b into pages of at most `room` points, starting a page at each manual break; a line bigger than a
 * page gets a page of its own. `titles` take room on every page but the ones that already show them.
 */
function bands(
  axis: Axis,
  a: number,
  b: number,
  room: number,
  breaks: readonly number[],
  titles: [number, number] | null,
): { from: number; to: number; titled: boolean }[] {
  const out: { from: number; to: number; titled: boolean }[] = []
  const titleSize = titles ? span(axis, titles[0], titles[1]) : 0
  let from = a
  while (from <= b) {
    // Titles repeat on pages that start after them.
    const titled = titles !== null && from > titles[1]
    const space = room - (titled ? titleSize : 0)
    let used = 0
    let to = from
    for (let i = from; i <= b; i++) {
      if (i > from && breaks.includes(i)) break
      const s = size(axis, i)
      if (i > from && used + s > space + 1e-6) break
      used += s
      to = i
    }
    out.push({ from, to, titled })
    from = to + 1
  }
  // Pages of hidden lines only print nothing.
  return out.filter((band) => span(axis, band.from, band.to) > 0)
}

/** The sheet's printed pages with its page setup. */
export function paginate(sheet: Sheet, setup: PageSetup, options: PaginateOptions = {}): Pagination {
  const inches = paperInches(setup, options.defaultPaper)
  const paper = { width: inches.width * 72, height: inches.height * 72 }
  const m = setup.margins
  const body = {
    x: m.left * 72,
    y: m.top * 72,
    width: Math.max(36, paper.width - (m.left + m.right) * 72),
    height: Math.max(36, paper.height - (m.top + m.bottom) * 72),
  }
  const used = printedRange(sheet)
  const areas = options.ranges ?? (!options.ignorePrintArea && setup.printArea ? setup.printArea : used ? [used] : [])
  const headings = setup.headings ? (options.headings ?? { width: 27, height: 15 }) : { width: 0, height: 0 }
  // Headings scale with the cells; the room left for cells at a scale.
  const roomAt = (scale: number) => ({
    width: body.width / scale - headings.width,
    height: body.height / scale - headings.height,
  })
  // Excel ignores manual page breaks while the sheet fits to pages.
  const rowBreaks = setup.fitTo ? [] : setup.rowBreaks
  const colBreaks = setup.fitTo ? [] : setup.colBreaks
  const layout = (scale: number) => {
    const room = roomAt(scale)
    return areas.map((area) => ({
      rows: bands(sheet.rows, area.r0, area.r1, room.height, rowBreaks, setup.titleRows),
      cols: bands(sheet.cols, area.c0, area.c1, room.width, colBreaks, setup.titleCols),
    }))
  }
  let scale = setup.scale / 100
  let parts = layout(scale)
  if (setup.fitTo) {
    // Fit to: the largest whole percent, at most 100%, at which the pages wide and tall fit (0: any number).
    const { width: wide, height: tall } = setup.fitTo
    const fits = (p: ReturnType<typeof layout>) =>
      p.every((a) => (!wide || a.cols.length <= wide) && (!tall || a.rows.length <= tall))
    scale = 0.1
    for (let percent = 100; percent >= 10; percent--) {
      const p = layout(percent / 100)
      if (fits(p) || percent === 10) {
        scale = percent / 100
        parts = p
        break
      }
    }
  }
  const pages: PrintPage[] = []
  let number = setup.firstPageNumber ?? options.firstNumber ?? 1
  parts.forEach((part, area) => {
    const page = (r: (typeof part.rows)[number], c: (typeof part.cols)[number]) =>
      pages.push({
        area,
        rows: [r.from, r.to],
        cols: [c.from, c.to],
        titleRows: r.titled ? setup.titleRows : null,
        titleCols: c.titled ? setup.titleCols : null,
        number: number++,
      })
    if (setup.order === "downThenOver") for (const c of part.cols) for (const r of part.rows) page(r, c)
    else for (const r of part.rows) for (const c of part.cols) page(r, c)
  })
  return { pages, areas, scale, paper, body }
}

/** The automatic page breaks Excel draws on the sheet as dashed lines once it has paged it: first rows and columns. */
export function pageBreaks(p: Pagination): { rows: number[]; cols: number[] } {
  const rows = new Set<number>()
  const cols = new Set<number>()
  for (const page of p.pages) {
    const area = p.areas[page.area]
    // Excel's dashed lines run between pages and along the far edge of the last ones.
    if (page.rows[0] > area.r0) rows.add(page.rows[0])
    if (page.cols[0] > area.c0) cols.add(page.cols[0])
    rows.add(page.rows[1] + 1)
    cols.add(page.cols[1] + 1)
  }
  return { rows: [...rows].sort((a, b) => a - b), cols: [...cols].sort((a, b) => a - b) }
}
