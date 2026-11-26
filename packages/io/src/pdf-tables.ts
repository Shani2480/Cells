/**
 * Data › Get Data › From PDF: the tables on a PDF's pages, found from where its text sits and the lines drawn around
 * it, as Power Query's Pdf.Tables lists them: each table as "Table001 (Page 1)", and each page's text in rows and
 * columns as "Page001". The input is the text on each page (from the PDF itself, or recognized on a scanned page) and
 * its ruling lines, in points with the origin at the page's top left; nothing here reads PDF files.
 *
 * Tables with lines around their cells take their grid from the lines. Tables without take their rows from the text's
 * lines and their columns from the gaps that run down every row. A table that goes on at the top of the next page,
 * with the same columns, is one table ("Table001 (Page 1-2)"), its repeated header row dropped.
 */

/**
 * A run of text on a page: a piece of a line in one font, or a chunk of recognized text. A recognized chunk gives
 * each word's place (`words`); its words stay one cell unless a column edge runs between them.
 */
export type PdfRun = {
  text: string
  x0: number
  y0: number
  x1: number
  y1: number
  size: number
  /** A recognized chunk's words; `unsure`: recognition wasn't sure of it. */
  words?: { text: string; x0: number; x1: number; unsure?: boolean }[]
}

/** A line drawn on the page (a thin filled rectangle counts too). */
export type PdfRule = { x0: number; y0: number; x1: number; y1: number }

export type PdfPageText = { width: number; height: number; runs: PdfRun[]; rules: PdfRule[] }

export type PdfItem = {
  /** "Table001", "Page001": Pdf.Tables' Id. */
  id: string
  /** "Table001 (Page 1)", "Table002 (Page 2-3)", "Page001". */
  name: string
  kind: "Table" | "Page"
  /** The pages it's on, from 1. */
  pages: number[]
  /** Its cells as text ("" for an empty cell), every row as wide as the widest. */
  rows: string[][]
  /** Cells holding recognized text that recognition wasn't sure of, as [row, column], for a person to check. */
  unsure?: [number, number][]
}

/** A word, or a piece of a run set as one; `chunk` is shared by a recognized chunk's words. */
type Word = {
  text: string
  x0: number
  y0: number
  x1: number
  y1: number
  size: number
  chunk?: number
  unsure?: boolean
}
type Line = { words: Word[]; y0: number; y1: number }
/** A table's cells, each the words in it. */
type Grid = Word[][][]
type Found = { page: number; top: number; bottom: number; cells: Grid }

const mid = (a: number, b: number) => (a + b) / 2

function median(values: number[]): number {
  if (!values.length) return 0
  const s = [...values].sort((a, b) => a - b)
  return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2
}

/**
 * A run split into its pieces where it has two or more spaces (or a tab) in a row, each placed by its share of the
 * run's characters. A single space keeps words together: a run is set as one piece of text, so its words never part
 * into cells.
 */
let chunks = 0

function wordsOf(run: PdfRun): Word[] {
  if (run.words?.length) {
    const chunk = ++chunks
    return run.words
      .filter((w) => w.text.trim())
      .map((w) => ({
        text: w.text.trim(),
        x0: w.x0,
        x1: w.x1,
        y0: run.y0,
        y1: run.y1,
        size: run.size,
        chunk,
        ...(w.unsure ? { unsure: true } : {}),
      }))
  }
  if (!run.text.trim()) return []
  const per = (run.x1 - run.x0) / Math.max(1, run.text.length)
  const out: Word[] = []
  for (const m of run.text.matchAll(/\S+(?: \S+)*/g)) {
    const at = m.index ?? 0
    out.push({
      text: m[0],
      x0: run.x0 + at * per,
      x1: run.x0 + (at + m[0].length) * per,
      y0: run.y0,
      y1: run.y1,
      size: run.size,
    })
  }
  return out
}

/** Words grouped into lines: a word joins the line it overlaps by half its height or more. */
function linesOf(words: Word[]): Line[] {
  const lines: Line[] = []
  for (const w of [...words].sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0)) {
    const h = w.y1 - w.y0
    const line = lines.find((l) => Math.min(l.y1, w.y1) - Math.max(l.y0, w.y0) >= 0.5 * Math.min(h, l.y1 - l.y0))
    if (line) {
      line.words.push(w)
      line.y0 = Math.min(line.y0, w.y0)
      line.y1 = Math.max(line.y1, w.y1)
    } else lines.push({ words: [w], y0: w.y0, y1: w.y1 })
  }
  for (const l of lines) l.words.sort((a, b) => a.x0 - b.x0)
  return lines.sort((a, b) => a.y0 - b.y0)
}

/** The gap between words that parts two cells: half the text's size (a space is about a quarter). */
const cellGap = (w: Word) => 0.5 * w.size

/** A line's phrases: words closer than a cell gap belong together. */
function phrases(line: Line): Word[][] {
  const out: Word[][] = []
  for (const w of line.words) {
    const last = out[out.length - 1]
    const prev = last?.[last.length - 1]
    if (prev && ((w.chunk !== undefined && w.chunk === prev.chunk) || w.x0 - prev.x1 < cellGap(w))) last.push(w)
    else out.push([w])
  }
  return out
}

const join = (words: Word[]) => words.map((w) => w.text).join(" ")

/** The text of a cell's words, left to right, its lines apart as Excel's Pdf.Tables gives them (a line break). */
function cellText(words: Word[]): string {
  return linesOf(words)
    .map((l) => join(l.words))
    .join("\n")
}

// ── Tables with lines ───────────────────────────────────────────────────────────────────────────────────────────

type Seg = { at: number; from: number; to: number }

/** Lines along one direction merged: the same position (within 1.5 pt), touching or overlapping. */
function mergeSegs(segs: Seg[]): Seg[] {
  const out: Seg[] = []
  for (const s of [...segs].sort((a, b) => a.at - b.at || a.from - b.from)) {
    const same = out.find((o) => Math.abs(o.at - s.at) <= 1.5 && s.from <= o.to + 2 && s.to >= o.from - 2)
    if (same) {
      same.from = Math.min(same.from, s.from)
      same.to = Math.max(same.to, s.to)
    } else out.push({ ...s })
  }
  return out
}

/** Positions within `tol` of each other as one (their middle). */
function cluster(values: number[], tol: number): number[] {
  const out: number[][] = []
  for (const v of [...values].sort((a, b) => a - b)) {
    const last = out[out.length - 1]
    if (last && v - last[last.length - 1] <= tol) last.push(v)
    else out.push([v])
  }
  return out.map((g) => mid(g[0], g[g.length - 1]))
}

/** The grids the page's lines draw: sets of crossing lines, each with two or more lines each way. */
function grids(rules: PdfRule[]): { xs: number[]; ys: number[] }[] {
  const hs: Seg[] = []
  const vs: Seg[] = []
  for (const r of rules) {
    const w = Math.abs(r.x1 - r.x0)
    const h = Math.abs(r.y1 - r.y0)
    const [x0, x1] = [Math.min(r.x0, r.x1), Math.max(r.x0, r.x1)]
    const [y0, y1] = [Math.min(r.y0, r.y1), Math.max(r.y0, r.y1)]
    if (h <= 3 && w > 3) hs.push({ at: mid(y0, y1), from: x0, to: x1 })
    else if (w <= 3 && h > 3) vs.push({ at: mid(x0, x1), from: y0, to: y1 })
  }
  const H = mergeSegs(hs)
  const V = mergeSegs(vs)
  // Crossing lines in one set (union-find over H then V).
  const parent = [...H, ...V].map((_, i) => i)
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])))
  H.forEach((h, i) =>
    V.forEach((v, j) => {
      if (v.at >= h.from - 3 && v.at <= h.to + 3 && h.at >= v.from - 3 && h.at <= v.to + 3)
        parent[find(i)] = find(H.length + j)
    }),
  )
  const sets = new Map<number, { h: Seg[]; v: Seg[] }>()
  ;[...H, ...V].forEach((s, i) => {
    const root = find(i)
    const set = sets.get(root) ?? { h: [], v: [] }
    if (i < H.length) set.h.push(s)
    else set.v.push(s)
    sets.set(root, set)
  })
  const out: { xs: number[]; ys: number[] }[] = []
  for (const { h, v } of sets.values()) {
    if (h.length < 2 || v.length < 2) continue
    const xs = cluster(
      v.map((s) => s.at),
      2,
    )
    const ys = cluster(
      h.map((s) => s.at),
      2,
    )
    if (xs.length >= 2 && ys.length >= 2) out.push({ xs, ys })
  }
  return out
}

/**
 * Rows and columns that hold no text taken out, and a heading that stands apart from its column (a column with only
 * a heading, next to one with no heading) put over it.
 */
function tidy(cells: Grid): Grid {
  let grid = cells.filter((r) => r.some((c) => c.length))
  const width = Math.max(0, ...grid.map((r) => r.length))
  const cols = Array.from({ length: width }, (_, i) => i).filter((i) => grid.some((r) => r[i]?.length))
  grid = grid.map((r) => cols.map((i) => r[i] ?? []))
  if (grid.length < 3) return grid
  for (let c = 0; c < (grid[0]?.length ?? 0); c++) {
    const onlyHead = grid[0][c].length > 0 && grid.slice(1).every((r) => !r[c].length)
    if (!onlyHead) continue
    const to = [c + 1, c - 1].find(
      (k) => k >= 0 && k < grid[0].length && !grid[0][k].length && grid.slice(1).some((r) => r[k].length),
    )
    if (to === undefined) continue
    grid[0][to] = grid[0][c]
    grid = grid.map((r) => r.filter((_, k) => k !== c))
    c--
  }
  return grid
}

function ruledTables(page: number, words: Word[], rules: PdfRule[], used: Set<Word>): Found[] {
  const out: Found[] = []
  for (const { xs, ys } of grids(rules)) {
    const cells: Word[][][] = ys.slice(1).map(() => xs.slice(1).map(() => []))
    let any = false
    for (const w of words) {
      if (used.has(w)) continue
      const cx = mid(w.x0, w.x1)
      const cy = mid(w.y0, w.y1)
      const r = ys.findIndex((y, i) => i < ys.length - 1 && cy >= y && cy < ys[i + 1])
      const c = xs.findIndex((x, i) => i < xs.length - 1 && cx >= x && cx < xs[i + 1])
      if (r < 0 || c < 0) continue
      cells[r][c].push(w)
      any = true
    }
    if (!any) continue
    const grid = tidy(cells)
    // A box around a paragraph, or one row of boxes, isn't a table.
    if (grid.length < 2 || (grid[0]?.length ?? 0) < 2) continue
    for (const row of cells) for (const cell of row) for (const w of cell) used.add(w)
    out.push({ page, top: ys[0], bottom: ys[ys.length - 1], cells: grid })
  }
  return out
}

// ── Tables without lines ────────────────────────────────────────────────────────────────────────────────────────

/**
 * The column edges of a set of lines: the middles of the gaps that run down them. A gap may be crossed by one line in
 * ten (a heading or a long value that runs over). A gap half the text's size or wider parts two columns; a narrower
 * one (down to a space) only where the text beside it lines up, as numbers set to the right of their column do, or
 * text set to the left of its column, when the columns almost touch.
 */
function columnEdges(lines: Line[]): number[] {
  const words = lines.flatMap((l) => l.words)
  if (!words.length) return []
  const left = Math.min(...words.map((w) => w.x0))
  const right = Math.max(...words.map((w) => w.x1))
  const step = 0.5
  const n = Math.ceil((right - left) / step) + 1
  const cover = new Array<number>(n).fill(0)
  for (const l of lines) {
    const hit = new Uint8Array(n)
    for (const w of l.words)
      for (let i = Math.floor((w.x0 - left) / step); i <= Math.ceil((w.x1 - left) / step) && i < n; i++) hit[i] = 1
    for (let i = 0; i < n; i++) cover[i] += hit[i]
  }
  const allowed = lines.length >= 10 ? Math.floor(lines.length / 10) : 0
  const size = median(words.map((w) => w.size))
  const edges: number[] = []
  let start = -1
  for (let i = 0; i <= n; i++) {
    const open = i < n && cover[i] <= allowed
    if (open && start < 0) start = i
    if (!open && start >= 0) {
      const from = left + start * step
      const to = left + i * step
      if (start > 0 && i < n && (to - from >= 0.5 * size || (to - from >= 0.2 * size && aligned(lines, from, to))))
        edges.push((from + to) / 2)
      start = -1
    }
  }
  return edges
}

/** Whether most lines have a word ending at a gap's left side, or most start one at its right side. */
function aligned(lines: Line[], from: number, to: number): boolean {
  let both = 0
  let ends = 0
  let starts = 0
  for (const l of lines) {
    const before = l.words.filter((w) => w.x1 <= from + 1).pop()
    const after = l.words.find((w) => w.x0 >= to - 1)
    if (!before || !after) continue
    both++
    if (Math.abs(before.x1 - from) <= 1.5) ends++
    if (Math.abs(after.x0 - to) <= 1.5) starts++
  }
  return both >= 0.6 * lines.length && Math.max(ends, starts) >= 0.6 * both
}

/** A line's words split at the column edges (a word goes by its middle). */
function splitLine(line: Line, edges: number[]): Word[][] {
  const cells: Word[][] = Array.from({ length: edges.length + 1 }, () => [])
  for (const w of line.words) {
    const cx = mid(w.x0, w.x1)
    let c = 0
    while (c < edges.length && cx > edges[c]) c++
    cells[c].push(w)
  }
  return cells
}

function unruledTables(page: number, words: Word[]): Found[] {
  const lines = linesOf(words)
  if (!lines.length) return []
  const height = median(lines.map((l) => l.y1 - l.y0))
  const multi = lines.map((l) => phrases(l).length >= 2)
  const out: Found[] = []
  let i = 0
  while (i < lines.length) {
    if (!multi[i]) {
      i++
      continue
    }
    // A block: lines of two or more cells, no more than three lines apart (rows can be spaced out); a one-cell line
    // inside (a wrapped value) may stay.
    const near = (j: number) => lines[j].y0 - lines[j - 1].y1 <= 3 * height
    let end = i
    for (let j = i + 1; j < lines.length; j++) {
      if (!near(j)) break
      if (multi[j]) end = j
      else if (!(j + 1 < lines.length && multi[j + 1] && near(j + 1))) break
    }
    const block = lines.slice(i, end + 1)
    i = end + 1
    const shaped = block.filter((l) => phrases(l).length >= 2)
    if (shaped.length < 3 && !(shaped.length === 2 && shaped.every((l) => phrases(l).length >= 3))) continue
    const edges = columnEdges(shaped)
    if (!edges.length) continue
    const rows = tidy(mergeWrapped(block, edges))
    // Most rows fill two or more cells, or it's text in columns, not a table.
    if (
      (rows[0]?.length ?? 0) < 2 ||
      rows.filter((r) => r.filter((c) => c.length).length >= 2).length < 0.6 * rows.length
    )
      continue
    out.push({ page, top: block[0].y0, bottom: block[block.length - 1].y1, cells: rows })
  }
  return out
}

/**
 * The block's rows: a line each, except a value wrapped onto more lines. A line with nothing in the first column and
 * at most half its cells filled goes with the row it sits closest to, above or below (cells are often set in the
 * middle of their row), when it's closer to it than the rows are to each other.
 */
function mergeWrapped(block: Line[], edges: number[]): Grid {
  const split = block.map((l) => splitLine(l, edges))
  const part = split.map(
    (cells) => !cells[0].length && cells.filter((c) => c.length).length <= Math.ceil(cells.length / 2),
  )
  const anchors = block.map((_, k) => k).filter((k) => !part[k])
  if (!anchors.length) return split
  const pitch = median(anchors.slice(1).map((k, n) => block[k].y0 - block[anchors[n]].y1))
  const owner = block.map((_, k) => k)
  block.forEach((line, k) => {
    if (!part[k]) return
    const above = [...anchors].reverse().find((a) => a < k)
    const below = anchors.find((a) => a > k)
    const up = above === undefined ? Infinity : line.y0 - block[above].y1
    const down = below === undefined ? Infinity : block[below].y0 - line.y1
    const best = up <= down ? above : below
    const gap = Math.min(up, down)
    if (best !== undefined && (anchors.length < 2 || gap < 0.8 * pitch)) owner[k] = best
  })
  const rows = new Map<number, Word[][]>()
  block.forEach((_, k) => {
    const row = rows.get(owner[k])
    if (row) split[k].forEach((c, i) => row[i].push(...c))
    else
      rows.set(
        owner[k],
        split[k].map((c) => [...c]),
      )
  })
  return [...rows.keys()].sort((a, b) => a - b).map((k) => rows.get(k) as Word[][])
}

// ── Pages ───────────────────────────────────────────────────────────────────────────────────────────────────────

const texts = (grid: Grid) => grid.map((r) => r.map(cellText))
const doubts = (grid: Grid) => grid.map((r) => r.map((c) => c.some((w) => w.unsure)))

/** The cells marked unsure, as [row, column]; undefined when there are none. */
function unsureCells(flags: boolean[][]): [number, number][] | undefined {
  const out: [number, number][] = []
  flags.forEach((r, i) => r.forEach((f, k) => f && out.push([i, k])))
  return out.length ? out : undefined
}

/** Where each column's text runs across the page. */
function spans(grid: Grid): [number, number][] {
  return (grid[0] ?? []).map((_, c) => {
    const ws = grid.flatMap((r) => r[c] ?? [])
    return [Math.min(...ws.map((w) => w.x0)), Math.max(...ws.map((w) => w.x1))]
  })
}

/**
 * Two tables that are one, carried over to the next page: the first ends low on its page, the second starts high on
 * the next, with as many columns, each running over its column in the first.
 */
function continues(a: Found, b: Found, pages: PdfPageText[]): boolean {
  if (b.page !== a.page + 1 || (a.cells[0]?.length ?? 0) !== (b.cells[0]?.length ?? 0)) return false
  if (a.bottom < 0.6 * pages[a.page].height || b.top > 0.4 * pages[b.page].height) return false
  const sa = spans(a.cells)
  const sb = spans(b.cells)
  return sa.every(([x0, x1], i) => Math.min(x1, sb[i][1]) > Math.max(x0, sb[i][0]))
}

const pad3 = (n: number) => String(n).padStart(3, "0")

function widen(rows: string[][]): string[][] {
  const width = Math.max(0, ...rows.map((r) => r.length))
  return rows.map((r) => [...r, ...new Array<string>(width - r.length).fill("")])
}

/** The PDF's tables and pages, as Pdf.Tables lists them (tables first, then pages). */
export function pdfTables(pages: PdfPageText[]): PdfItem[] {
  const found: Found[] = []
  const pageItems: PdfItem[] = []
  pages.forEach((p, index) => {
    const words = p.runs.flatMap(wordsOf)
    // The page: a row per line, columns from the gaps in its lines of two or more cells.
    const lines = linesOf(words)
    const pageGrid = tidy(lines.map((l) => splitLine(l, columnEdges(lines.filter((x) => phrases(x).length >= 2)))))
    const used = new Set<Word>()
    const ruled = ruledTables(index, words, p.rules, used)
    const loose = unruledTables(
      index,
      words.filter((w) => !used.has(w)),
    )
    found.push(...[...ruled, ...loose].sort((a, b) => a.top - b.top))
    const unsure = unsureCells(doubts(pageGrid))
    pageItems.push({
      id: `Page${pad3(index + 1)}`,
      name: `Page${pad3(index + 1)}`,
      kind: "Page",
      pages: [index + 1],
      rows: widen(texts(pageGrid)),
      ...(unsure ? { unsure } : {}),
    })
  })
  // Tables carried over to the next page joined, a repeated header row dropped.
  const joined: { pages: number[]; rows: string[][]; flags: boolean[][]; last: Found }[] = []
  for (const t of found) {
    const prev = joined[joined.length - 1]
    const rows = texts(t.cells)
    const flags = doubts(t.cells)
    if (prev && continues(prev.last, t, pages)) {
      const skip = rows[0]?.join("\t") === prev.rows[0]?.join("\t") ? 1 : 0
      prev.rows.push(...rows.slice(skip))
      prev.flags.push(...flags.slice(skip))
      prev.pages.push(t.page + 1)
      prev.last = t
    } else joined.push({ pages: [t.page + 1], rows, flags, last: t })
  }
  const tables = joined.map((t, i): PdfItem => {
    const first = t.pages[0]
    const last = t.pages[t.pages.length - 1]
    const id = `Table${pad3(i + 1)}`
    const unsure = unsureCells(t.flags)
    return {
      id,
      name: `${id} (Page ${first === last ? first : `${first}-${last}`})`,
      kind: "Table",
      pages: t.pages,
      rows: widen(t.rows),
      ...(unsure ? { unsure } : {}),
    }
  })
  return [...tables, ...pageItems]
}
