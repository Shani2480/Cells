import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { CellValue } from "./values"

/** Everything stored for one cell. A cell with only a style is "formatted but empty". */
export type CellData = {
  value?: CellValue
  /** Typed text when it differs from the value's own text (formulas, a leading apostrophe). */
  input?: string
  /** Style id in the workbook's StyleTable; absent means 0 (default). */
  style?: number
  /** The value was spilled here by a dynamic array formula (it is not the cell's own content). */
  spill?: true
}

export type CellEntry = { row: number; col: number; data: CellData }

const CHUNK_BITS = 10
/** Rows per chunk. Columns are stored as chunks of this many rows, allocated on first write. */
export const CHUNK_SIZE = 1 << CHUNK_BITS
const CHUNK_MASK = CHUNK_SIZE - 1

class Chunk {
  readonly values: (CellValue | undefined)[] = new Array(CHUNK_SIZE)
  inputs: Map<number, string> | null = null
  styles: Uint32Array | null = null
  /** Offsets whose value was spilled by an array formula. */
  spills: Set<number> | null = null
  /** Cells with a value. */
  valueCount = 0
  /** Cells with a value or a style. */
  storedCount = 0

  has(offset: number) {
    return this.values[offset] !== undefined || (this.styles !== null && this.styles[offset] !== 0)
  }
}

class Column {
  readonly chunks = new Map<number, Chunk>()
  valueCount = 0
  storedCount = 0
}

/**
 * Sparse cell storage for one sheet: column → chunk of 1,024 rows → parallel arrays. Empty chunks and columns are
 * dropped, so scans skip empty space in chunk-sized steps.
 */
export class CellStore {
  private readonly columns = new Map<number, Column>()
  private sortedColumns: number[] | null = null
  private boundsCache: Range | null | undefined = undefined

  get(row: number, col: number): CellData | undefined {
    const chunk = this.chunk(row, col)
    if (!chunk) {
      return undefined
    }
    const offset = row & CHUNK_MASK
    if (!chunk.has(offset)) {
      return undefined
    }
    const data: CellData = {}
    const value = chunk.values[offset]
    if (value !== undefined) data.value = value
    const input = chunk.inputs?.get(offset)
    if (input !== undefined) data.input = input
    const style = chunk.styles?.[offset] ?? 0
    if (style !== 0) data.style = style
    if (chunk.spills?.has(offset)) data.spill = true
    return data
  }

  /** True when the cell's value was spilled by an array formula. */
  isSpill(row: number, col: number): boolean {
    return this.chunk(row, col)?.spills?.has(row & CHUNK_MASK) ?? false
  }

  /** A formula's calculated value: replaces the value only, keeping the formula text and the style. */
  setComputed(row: number, col: number, value: CellValue | undefined) {
    const chunk = this.chunk(row, col)
    const offset = row & CHUNK_MASK
    this.put(row, col, value, chunk?.inputs?.get(offset), chunk?.styles?.[offset] ?? 0, false)
  }

  /** A spilled value (or undefined to take it away), keeping the cell's style. */
  setSpill(row: number, col: number, value: CellValue | undefined) {
    const chunk = this.chunk(row, col)
    const style = chunk?.styles?.[row & CHUNK_MASK] ?? 0
    this.put(row, col, value, undefined, style, value !== undefined)
  }

  /** Visits every stored formula (input starting with "="). */
  forEachFormula(visit: (row: number, col: number, input: string) => void) {
    for (const [col, column] of this.columns) {
      for (const [key, chunk] of column.chunks) {
        if (!chunk.inputs) continue
        for (const [offset, input] of chunk.inputs) {
          if (input.startsWith("=")) visit((key << CHUNK_BITS) + offset, col, input)
        }
      }
    }
  }

  value(row: number, col: number): CellValue | undefined {
    return this.chunk(row, col)?.values[row & CHUNK_MASK]
  }

  hasValue(row: number, col: number): boolean {
    return this.value(row, col) !== undefined
  }

  input(row: number, col: number): string | undefined {
    return this.chunk(row, col)?.inputs?.get(row & CHUNK_MASK)
  }

  style(row: number, col: number): number {
    return this.chunk(row, col)?.styles?.[row & CHUNK_MASK] ?? 0
  }

  /** Sets only the cell's style id (0 removes it), keeping its value and formula. */
  setStyle(row: number, col: number, style: number) {
    const chunk = this.chunk(row, col)
    const offset = row & CHUNK_MASK
    if (!chunk && style === 0) return
    this.put(row, col, chunk?.values[offset], chunk?.inputs?.get(offset), style, chunk?.spills?.has(offset) ?? false)
  }

  /** Visits the stored cells in `range` with their style ids, without building a record per cell. */
  forEachStyle(range: Range, visit: (row: number, col: number, style: number) => void) {
    for (const col of this.usedColumns()) {
      if (col < range.c0) continue
      if (col > range.c1) break
      const column = this.columns.get(col)!
      for (const [key, chunk] of column.chunks) {
        const base = key << CHUNK_BITS
        if (base > range.r1 || base + CHUNK_MASK < range.r0) continue
        const start = Math.max(range.r0, base) - base
        const end = Math.min(range.r1, base + CHUNK_MASK) - base
        const styles = chunk.styles
        for (let offset = start; offset <= end; offset++) {
          if (chunk.has(offset)) visit(base + offset, col, styles ? styles[offset] : 0)
        }
      }
    }
  }

  /** Replaces everything stored for the cell; `undefined` or an empty record removes it. */
  set(row: number, col: number, data: CellData | undefined) {
    this.put(
      row,
      col,
      data?.value,
      data?.value !== undefined ? data?.input : undefined,
      data?.style ?? 0,
      data?.spill === true,
    )
  }

  private put(
    row: number,
    col: number,
    value: CellValue | undefined,
    input: string | undefined,
    style: number,
    spill: boolean,
  ) {
    const stored = value !== undefined || style !== 0
    let column = this.columns.get(col)
    if (!column) {
      if (!stored) return
      column = new Column()
      this.columns.set(col, column)
      this.sortedColumns = null
    }
    const key = row >>> CHUNK_BITS
    let chunk = column.chunks.get(key)
    if (!chunk) {
      if (!stored) return
      chunk = new Chunk()
      column.chunks.set(key, chunk)
    }
    const offset = row & CHUNK_MASK
    const hadValue = chunk.values[offset] !== undefined
    const wasStored = chunk.has(offset)

    chunk.values[offset] = value
    if (input !== undefined) {
      chunk.inputs ??= new Map()
      chunk.inputs.set(offset, input)
    } else if (chunk.inputs) {
      chunk.inputs.delete(offset)
    }
    if (style !== 0) {
      chunk.styles ??= new Uint32Array(CHUNK_SIZE)
      chunk.styles[offset] = style
    } else if (chunk.styles) {
      chunk.styles[offset] = 0
    }
    if (spill && value !== undefined) (chunk.spills ??= new Set()).add(offset)
    else chunk.spills?.delete(offset)

    const valueDelta = (value !== undefined ? 1 : 0) - (hadValue ? 1 : 0)
    const storedDelta = (stored ? 1 : 0) - (wasStored ? 1 : 0)
    chunk.valueCount += valueDelta
    chunk.storedCount += storedDelta
    column.valueCount += valueDelta
    column.storedCount += storedDelta
    if (storedDelta !== 0) {
      this.boundsCache = undefined
    }
    if (chunk.storedCount === 0) {
      column.chunks.delete(key)
    }
    if (column.storedCount === 0) {
      this.columns.delete(col)
      this.sortedColumns = null
    }
  }

  /** Columns that store anything, ascending. */
  usedColumns(): readonly number[] {
    this.sortedColumns ??= [...this.columns.keys()].sort((a, b) => a - b)
    return this.sortedColumns
  }

  columnHasValues(col: number) {
    return (this.columns.get(col)?.valueCount ?? 0) > 0
  }

  /**
   * The nearest row after `from` (dir 1) or before it (dir -1) in `col` that holds a value, or -1. Empty chunks are
   * skipped whole.
   */
  nextValueRow(col: number, from: number, dir: 1 | -1): number {
    const column = this.columns.get(col)
    if (!column || column.valueCount === 0) {
      return -1
    }
    let row = from + dir
    while (row >= 0 && row < MAX_ROWS) {
      const key = row >>> CHUNK_BITS
      const chunk = column.chunks.get(key)
      if (!chunk || chunk.valueCount === 0) {
        const next = this.nextChunkKey(column, key, dir)
        if (next === -1) {
          return -1
        }
        row = dir === 1 ? next << CHUNK_BITS : (next << CHUNK_BITS) + CHUNK_MASK
        continue
      }
      const end = dir === 1 ? (key << CHUNK_BITS) + CHUNK_MASK : key << CHUNK_BITS
      for (; dir === 1 ? row <= end : row >= end; row += dir) {
        if (chunk.values[row & CHUNK_MASK] !== undefined) {
          return row
        }
      }
    }
    return -1
  }

  private nextChunkKey(column: Column, key: number, dir: 1 | -1): number {
    let best = -1
    for (const [other, chunk] of column.chunks) {
      if (chunk.valueCount === 0) continue
      if (dir === 1 ? other > key && (best === -1 || other < best) : other < key && other > best) {
        best = other
      }
    }
    return best
  }

  /** Visits stored cells inside `range`, column by column, rows ascending. */
  forEach(range: Range, visit: (row: number, col: number, data: CellData) => void, onlyValues = false) {
    for (const col of this.usedColumns()) {
      if (col < range.c0) continue
      if (col > range.c1) break
      const column = this.columns.get(col)!
      const keys = [...column.chunks.keys()].filter(
        (key) => key >= range.r0 >>> CHUNK_BITS && key <= range.r1 >>> CHUNK_BITS,
      )
      keys.sort((a, b) => a - b)
      for (const key of keys) {
        const chunk = column.chunks.get(key)!
        if (onlyValues && chunk.valueCount === 0) continue
        const base = key << CHUNK_BITS
        const start = Math.max(range.r0, base) - base
        const end = Math.min(range.r1, base + CHUNK_MASK) - base
        for (let offset = start; offset <= end; offset++) {
          if (onlyValues ? chunk.values[offset] !== undefined : chunk.has(offset)) {
            visit(base + offset, col, this.get(base + offset, col)!)
          }
        }
      }
    }
  }

  /** All stored cells inside `range`, for undo records. */
  entries(range: Range): CellEntry[] {
    const out: CellEntry[] = []
    this.forEach(range, (row, col, data) => out.push({ row, col, data }))
    return out
  }

  /** The smallest range holding every stored cell (values or formats), or null for an empty sheet. */
  usedBounds(): Range | null {
    if (this.boundsCache !== undefined) {
      return this.boundsCache
    }
    const cols = this.usedColumns()
    if (cols.length === 0) {
      this.boundsCache = null
      return null
    }
    let r0 = MAX_ROWS
    let r1 = -1
    for (const col of cols) {
      for (const [key, chunk] of this.columns.get(col)!.chunks) {
        const base = key << CHUNK_BITS
        if (base < r0 || base + CHUNK_MASK > r1) {
          for (let offset = 0; offset < CHUNK_SIZE; offset++) {
            if (chunk.has(offset)) {
              r0 = Math.min(r0, base + offset)
              r1 = Math.max(r1, base + offset)
            }
          }
        }
      }
    }
    this.boundsCache = { r0, r1, c0: cols[0], c1: cols[cols.length - 1] }
    return this.boundsCache
  }

  /** Total cells with a value. */
  get valueCount() {
    let n = 0
    for (const column of this.columns.values()) n += column.valueCount
    return n
  }

  /** Clears every cell inside `range`. */
  clear(range: Range) {
    for (const entry of this.entries(range)) {
      this.set(entry.row, entry.col, undefined)
    }
  }

  /**
   * Moves the block of cells in columns c0..c1 and rows ≥ `from` by `delta` rows. Cells pushed past the last row are
   * dropped; the caller clears (and records) the rows a deletion removes before shifting up. Works on the chunk
   * arrays directly: no per-cell objects, so inserting a row above 100,000 filled rows stays fast.
   */
  shiftRows(from: number, delta: number, c0 = 0, c1 = MAX_COLS - 1) {
    if (delta === 0) return
    for (const col of [...this.usedColumns()]) {
      if (col < c0 || col > c1) continue
      const column = this.columns.get(col)!
      const rebuilt = new Map<number, Chunk>()
      let target: Chunk | undefined
      let targetKey = -1
      const write = (row: number, from: Chunk, offset: number) => {
        const key = row >>> CHUNK_BITS
        if (key !== targetKey) {
          targetKey = key
          target = rebuilt.get(key)
          if (!target) {
            target = new Chunk()
            rebuilt.set(key, target)
          }
        }
        const to = row & CHUNK_MASK
        const value = from.values[offset]
        target!.values[to] = value
        const input = from.inputs?.get(offset)
        if (input !== undefined) (target!.inputs ??= new Map()).set(to, input)
        const style = from.styles?.[offset] ?? 0
        if (style !== 0) (target!.styles ??= new Uint32Array(CHUNK_SIZE))[to] = style
        if (from.spills?.has(offset)) (target!.spills ??= new Set()).add(to)
        if (value !== undefined) target!.valueCount++
        target!.storedCount++
      }
      for (const [key, chunk] of column.chunks) {
        const base = key << CHUNK_BITS
        if (base + CHUNK_MASK < from) {
          // Untouched chunk above the shift: keep it as it is.
          rebuilt.set(key, chunk)
          continue
        }
        for (let offset = 0; offset < CHUNK_SIZE; offset++) {
          if (!chunk.has(offset)) continue
          const row = base + offset
          const moved = row < from ? row : row + delta
          if (moved >= 0 && moved < MAX_ROWS) write(moved, chunk, offset)
        }
      }
      column.chunks.clear()
      column.valueCount = 0
      column.storedCount = 0
      for (const [key, chunk] of rebuilt) {
        column.chunks.set(key, chunk)
        column.valueCount += chunk.valueCount
        column.storedCount += chunk.storedCount
      }
      if (column.storedCount === 0) {
        this.columns.delete(col)
        this.sortedColumns = null
      }
    }
    this.boundsCache = undefined
  }

  /** As shiftRows, across: columns ≥ `from` in rows r0..r1 move by `delta`. Whole columns move by re-keying. */
  shiftCols(from: number, delta: number, r0 = 0, r1 = MAX_ROWS - 1) {
    if (delta === 0) return
    const affected = this.usedColumns().filter((col) => col >= from)
    if (r0 === 0 && r1 === MAX_ROWS - 1) {
      const moving = affected.map((col) => [col, this.columns.get(col)!] as const)
      for (const [col] of moving) this.columns.delete(col)
      for (const [col, column] of moving) {
        const target = col + delta
        if (target >= 0 && target < MAX_COLS) this.columns.set(target, column)
      }
      this.sortedColumns = null
      this.boundsCache = undefined
      return
    }
    const moved = this.entries({ r0, r1, c0: from, c1: MAX_COLS - 1 })
    for (const entry of moved) this.set(entry.row, entry.col, undefined)
    for (const entry of moved) {
      const col = entry.col + delta
      if (col >= 0 && col < MAX_COLS) this.set(entry.row, col, entry.data)
    }
  }

  private chunk(row: number, col: number): Chunk | undefined {
    return this.columns.get(col)?.chunks.get(row >>> CHUNK_BITS)
  }
}
