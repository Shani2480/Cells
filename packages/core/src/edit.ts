import { MAX_COLS, MAX_ROWS, rangeCols, rangeRows, type Range } from "./address"
import type { CellData } from "./cell-store"
import { fillSeries } from "./fill"
import { copyFormula } from "./formula/adjust"
import { completeFormula } from "./formula/editing"
import type { CellWrite, Op } from "./ops"
import type { Sheet } from "./sheet"
import { DEFAULT_STYLE_ID, type StyleTable } from "./styles"
import { parseInput } from "./values"

/** A rectangle of cell contents, as copied. `cells[r][c]` is undefined for an empty cell. */
export type CellBlock = {
  rows: number
  cols: number
  cells: (CellData | undefined)[][]
  /** Where the block was copied from; pasted formulas move their relative references by the distance. */
  origin?: { row: number; col: number }
}

/** A cell's content written `dr` rows and `dc` columns from where it was: formulas adjust relative references. */
function moved(data: CellData, dr: number, dc: number): CellData {
  const copy = { ...data }
  delete copy.spill
  if (copy.input?.startsWith("=")) copy.input = copyFormula(copy.input, dr, dc)
  return copy
}

/**
 * The cell content for typed text, keeping the cell's format. A number typed as "10%" or "1,000" gives the cell
 * that number format when it has none (General), as Excel does on entry.
 */
export function contentForInput(
  styles: StyleTable,
  current: CellData | undefined,
  text: string,
  /** The row's or column's format, which a cell without its own shows. */
  inherited = DEFAULT_STYLE_ID,
): CellData | undefined {
  const style = current?.style ?? DEFAULT_STYLE_ID
  const parsed = parseInput(completeFormula(text))
  if (!parsed) {
    return style ? { style } : undefined
  }
  let nextStyle = style
  const shown = style || inherited
  if (parsed.numberFormat && !styles.get(shown).numberFormat) {
    // The cell gets its own format: what it showed (its row's or column's) plus the number format.
    nextStyle = styles.derive(shown, { numberFormat: parsed.numberFormat })
  }
  const data: CellData = { value: parsed.value }
  if (parsed.input !== undefined) data.input = parsed.input
  if (nextStyle) data.style = nextStyle
  return data
}

const lineStyle = (sheet: Sheet, row: number, col: number) => sheet.rowStyles.get(row) || sheet.colStyles.get(col)

/** Typing into one cell. */
export function writeInputOp(
  sheetIndex: number,
  sheet: Sheet,
  styles: StyleTable,
  row: number,
  col: number,
  text: string,
): Op {
  return {
    kind: "setCells",
    sheet: sheetIndex,
    cells: [{ row, col, data: contentForInput(styles, sheet.cells.get(row, col), text, lineStyle(sheet, row, col)) }],
  }
}

/**
 * Ctrl+Enter: the same text into every selected cell. A formula is written as typed in the active cell (`origin`)
 * and moves its relative references for the others, as if copied there.
 */
export function fillInputOp(
  sheetIndex: number,
  sheet: Sheet,
  styles: StyleTable,
  ranges: readonly Range[],
  text: string,
  origin?: { row: number; col: number },
): Op {
  const cells: CellWrite[] = []
  const formula = origin && text.startsWith("=")
  for (const range of ranges) {
    for (let row = range.r0; row <= range.r1; row++) {
      for (let col = range.c0; col <= range.c1; col++) {
        const typed = formula ? copyFormula(text, row - origin.row, col - origin.col) : text
        const data = contentForInput(styles, sheet.cells.get(row, col), typed, lineStyle(sheet, row, col))
        cells.push({ row, col, data })
      }
    }
  }
  return { kind: "setCells", sheet: sheetIndex, cells }
}

/** Delete: removes values (and formulas), keeps formats. Only stored cells are touched, so whole columns are cheap. */
export function clearContentsOp(sheetIndex: number, sheet: Sheet, ranges: readonly Range[]): Op | null {
  const cells: CellWrite[] = []
  for (const range of ranges) {
    sheet.cells.forEach(range, (row, col, data) => {
      if (data.value !== undefined) {
        cells.push({ row, col, data: data.style ? { style: data.style } : undefined })
      }
    })
  }
  return cells.length ? { kind: "setCells", sheet: sheetIndex, cells } : null
}

/** Copies a range's contents into a block. */
export function readBlock(sheet: Sheet, range: Range): CellBlock {
  const rows = rangeRows(range)
  const cols = rangeCols(range)
  const cells: (CellData | undefined)[][] = Array.from({ length: rows }, () => new Array(cols))
  sheet.cells.forEach(range, (row, col, data) => {
    cells[row - range.r0][col - range.c0] = data
  })
  return { rows, cols, cells, origin: { row: range.r0, col: range.c0 } }
}

/**
 * Where a paste lands: the block's own size from the target's top-left, or the whole target when the target is a
 * whole multiple of the block in both directions (Excel tiles the copy). Clamped to the sheet.
 */
export function pasteArea(target: Range, block: CellBlock): Range {
  const rows = rangeRows(target)
  const cols = rangeCols(target)
  const tiles = rows % block.rows === 0 && cols % block.cols === 0 && (rows > block.rows || cols > block.cols)
  const area = tiles
    ? target
    : { r0: target.r0, c0: target.c0, r1: target.r0 + block.rows - 1, c1: target.c0 + block.cols - 1 }
  return { ...area, r1: Math.min(MAX_ROWS - 1, area.r1), c1: Math.min(MAX_COLS - 1, area.c1) }
}

/** Writes `block` over `area` (from pasteArea), repeating it to fill. */
export function pasteBlockOp(sheetIndex: number, area: Range, block: CellBlock): Op {
  const cells: CellWrite[] = []
  for (let row = area.r0; row <= area.r1; row++) {
    const br = (row - area.r0) % block.rows
    const source = block.cells[br]
    for (let col = area.c0; col <= area.c1; col++) {
      const bc = (col - area.c0) % block.cols
      const data = source[bc]
      const origin = block.origin
      const dr = origin ? row - (origin.row + br) : 0
      const dc = origin ? col - (origin.col + bc) : 0
      cells.push({ row, col, data: data ? moved(data, dr, dc) : undefined })
    }
  }
  return { kind: "setCells", sheet: sheetIndex, cells }
}

/** Cut then paste: the source empties and its cells land at `area`'s top-left; references to them follow. */
export function moveBlockOps(sheetIndex: number, source: Range, area: Range): Op[] {
  const dr = area.r0 - source.r0
  const dc = area.c0 - source.c0
  if (dr === 0 && dc === 0) return []
  const r1 = Math.min(MAX_ROWS - 1, source.r1 + dr) - dr
  const c1 = Math.min(MAX_COLS - 1, source.c1 + dc) - dc
  return [{ kind: "moveCells", sheet: sheetIndex, source: { ...source, r1, c1 }, dr, dc }]
}

/** Plain text rows (from the clipboard) as a block of parsed cells. */
export function blockFromText(styles: StyleTable, rows: readonly (readonly string[])[]): CellBlock {
  const cols = Math.max(1, ...rows.map((row) => row.length))
  const cells = rows.map((row) => {
    const out: (CellData | undefined)[] = new Array(cols)
    row.forEach((text, col) => {
      out[col] = contentForInput(styles, undefined, text)
    })
    return out
  })
  return { rows: Math.max(1, rows.length), cols, cells: cells.length ? cells : [new Array(cols)] }
}

/**
 * The fill handle: `source` is the selection, `target` the selection grown by the drag (sharing a corner with it).
 * Each column (filling down/up) or row (filling right/left) continues its own series.
 */
export function fillOp(sheetIndex: number, sheet: Sheet, source: Range, target: Range, invert = false): Op | null {
  const cells: CellWrite[] = []
  const block = readBlock(sheet, source)
  if (target.r1 > source.r1 || target.r0 < source.r0) {
    const dir = target.r1 > source.r1 ? 1 : -1
    const count = dir === 1 ? target.r1 - source.r1 : source.r0 - target.r0
    for (let c = 0; c < block.cols; c++) {
      const column = block.cells.map((row) => row[c])
      fillSeries(column, count, dir, invert, (input, steps) => copyFormula(input, steps, 0)).forEach((data, k) => {
        const row = dir === 1 ? source.r1 + 1 + k : source.r0 - 1 - k
        cells.push({ row, col: source.c0 + c, data })
      })
    }
  } else if (target.c1 > source.c1 || target.c0 < source.c0) {
    const dir = target.c1 > source.c1 ? 1 : -1
    const count = dir === 1 ? target.c1 - source.c1 : source.c0 - target.c0
    for (let r = 0; r < block.rows; r++) {
      const row = Array.from({ length: block.cols }, (_, c) => block.cells[r][c])
      fillSeries(row, count, dir, invert, (input, steps) => copyFormula(input, 0, steps)).forEach((data, k) => {
        const col = dir === 1 ? source.c1 + 1 + k : source.c0 - 1 - k
        cells.push({ row: source.r0 + r, col, data })
      })
    }
  }
  return cells.length ? { kind: "setCells", sheet: sheetIndex, cells } : null
}
