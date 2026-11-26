import { pageBreaks, PT_PER_PX, type Axis, type Pagination } from "@cells/core"

/**
 * View › Page Layout: where each row and column sits when the sheet is shown as sheets of paper. The pages follow
 * one another with a gap between them, each with its margins around the cells it prints; past the printed cells
 * they go on, filled as printing would fill them. All sizes are sheet pixels at 100% (the print scale applied).
 *
 * `pos`, `indexAt` and `total` stand in for the axis's own offset, indexAt and total while Page Layout is shown.
 */

/** The space around and between the pages (sheet pixels). */
export const PAGE_GAP = 14

/** One page's run of rows (or columns) along an axis. */
export type PageBand = { first: number; last: number; start: number }

export type AxisMetric = {
  /** Where a line starts, and the line at a position (both sheet pixels from the start). */
  pos(index: number): number
  indexAt(px: number): number
  total: number
  /** The lines shown: those before this one. */
  count: number
}

export type PageAxis = AxisMetric & {
  bands: PageBand[]
  /** The page's size along the axis, and the margin before the cells (the left or top margin). */
  page: number
  margin: number
  /** The printable part's size (the page less both margins). */
  body: number
}

/** How many pages to carry on past the printed cells (Excel shows blank pages to type into). */
const EXTRA_PAGES = { rows: 4, cols: 2 }

/**
 * The rows or columns laid out on pages, from the sheet's pages (their breaks) and the paper. Without a print area
 * the last printed page fills up with the lines after it, as Excel's Page Layout does (a sheet with data in column A
 * shows A to H on its first page); a print area's far edge ends its pages.
 */
export function pageAxis(axis: Axis, p: Pagination, which: "rows" | "cols", used: number, printArea = false): PageAxis {
  const scale = p.scale || 1
  const toPx = (pt: number) => pt / PT_PER_PX / scale
  const page = toPx(which === "cols" ? p.paper.width : p.paper.height)
  const margin = toPx(which === "cols" ? p.body.x : p.body.y)
  const body = Math.max(1, toPx(which === "cols" ? p.body.width : p.body.height))
  const all = which === "cols" ? pageBreaks(p).cols : pageBreaks(p).rows
  // Where printed pages start (the far edges of the last ones only count for a print area).
  const starts = new Set(p.pages.map((page) => (which === "cols" ? page.cols[0] : page.rows[0])))
  const breaks = printArea ? all : all.filter((b) => starts.has(b))
  const firsts: number[] = []
  // Greedy pages from `from` while they start before `to`: as many lines as fit the printable part (at least one),
  // up to `limit`. Returns where the next page starts.
  const fill = (from: number, to: number, limit: number) => {
    let i = from
    while (i < to) {
      firsts.push(i)
      i = nextFirst(axis, i, body, limit)
    }
    return i
  }
  const points = [...new Set([0, ...breaks])].sort((a, b) => a - b)
  for (let k = 0; k < points.length - 1; k++) fill(points[k], points[k + 1], points[k + 1])
  const lastPoint = points[points.length - 1]
  // The last printed page, and the rest of the cells in use, as full pages.
  let end = fill(lastPoint, Math.min(axis.count, Math.max(lastPoint + 1, used + 1)), axis.count)
  // A few blank pages past the end.
  for (let extra = 0; extra < EXTRA_PAGES[which] && end < axis.count; extra++) {
    firsts.push(end)
    end = nextFirst(axis, end, body, axis.count)
  }
  const step = page + PAGE_GAP
  const bands: PageBand[] = firsts.map((first, i) => ({
    first,
    last: (i + 1 < firsts.length ? firsts[i + 1] : end) - 1,
    start: PAGE_GAP + i * step,
  }))
  const bandOf = (index: number) => {
    let lo = 0
    let hi = bands.length - 1
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (bands[mid].first <= index) lo = mid
      else hi = mid - 1
    }
    return bands[lo]
  }
  const last = bands[bands.length - 1]
  return {
    bands,
    page,
    margin,
    body,
    total: last.start + page + PAGE_GAP,
    count: end,
    pos(index) {
      const b = bandOf(index)
      return b.start + margin + axis.offset(index) - axis.offset(b.first)
    },
    indexAt(px) {
      const i = Math.max(0, Math.min(bands.length - 1, Math.floor((px - PAGE_GAP) / step)))
      const b = bands[i]
      const local = px - b.start - margin
      if (local <= 0) return b.first
      return Math.min(b.last, axis.indexAt(axis.offset(b.first) + local))
    },
  }
}

/** Where the page after one starting at `first` starts. */
function nextFirst(axis: Axis, first: number, body: number, count: number): number {
  let size = 0
  let j = first
  while (j < count && (j === first || size + axis.size(j) <= body)) size += axis.size(j++)
  return j
}

/** The axis as it is (Normal view and Page Break Preview). */
export function plainMetric(axis: Axis): AxisMetric {
  return { pos: (i) => axis.offset(i), indexAt: (px) => axis.indexAt(px), total: axis.total, count: axis.count }
}
