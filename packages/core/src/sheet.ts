import type { SheetChart } from "./chart"
import { defaultPageSetup, type PageSetup } from "./page-setup"
import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import { Axis } from "./axis"
import { CellStore } from "./cell-store"
import { Annotations } from "./annotations"
import type { CFEntry } from "./conditional"
import type { ValidationEntry } from "./validation"
import type { Table } from "./table/table"
import type { AutoFilter } from "./filter"
import type { SheetProtection } from "./protection"
import { emptyScenarios, type ScenarioSet } from "./scenarios"

/**
 * Default sizes at 100% zoom, as Excel draws a new workbook with an 11 pt body font: rows 15 pt (20 px), columns
 * 8.43 characters (64 px).
 */
export const DEFAULT_ROW_HEIGHT = 20
export const DEFAULT_COL_WIDTH = 64

/** Style ids of whole rows or columns (formatting all of column B), sparse. */
export class LineStyles {
  private readonly ids = new Map<number, number>()

  constructor(readonly count: number) {}

  get(index: number): number {
    return this.ids.get(index) ?? 0
  }

  set(index: number, id: number) {
    if (id === 0) this.ids.delete(index)
    else this.ids.set(index, id)
  }

  get size() {
    return this.ids.size
  }

  /** Lines with a style in [from, to], as [index, id] pairs. */
  entries(from = 0, to = this.count - 1): [number, number][] {
    const out: [number, number][] = []
    for (const [i, id] of this.ids) if (i >= from && i <= to) out.push([i, id])
    return out.sort((a, b) => a[0] - b[0])
  }

  /**
   * Inserts `n` lines at `at`. Like Excel's default "Format Same As Above/Left", they take the style of the line
   * before them. Lines pushed past the end fall off.
   */
  insert(at: number, n: number) {
    const moved = [...this.ids].map(([i, id]) => [i >= at ? i + n : i, id] as const)
    this.ids.clear()
    for (const [i, id] of moved) if (i < this.count) this.ids.set(i, id)
    const before = at > 0 ? this.get(at - 1) : 0
    if (before) for (let i = at; i < at + n && i < this.count; i++) this.ids.set(i, before)
  }

  /** Removes `n` lines at `at`. */
  remove(at: number, n: number) {
    const moved = [...this.ids].filter(([i]) => i < at || i >= at + n).map(([i, id]) => [i >= at ? i - n : i, id])
    this.ids.clear()
    for (const [i, id] of moved) this.ids.set(i, id)
  }
}

/** Rows per bucket in the merge index. */
const BUCKET = 64

/** The sheet's merged ranges, with a lookup by cell that stays fast with many merges. */
export class Merges {
  private readonly list: Range[] = []
  private buckets: Map<number, Range[]> | null = null

  get all(): readonly Range[] {
    return this.list
  }

  get size() {
    return this.list.length
  }

  /** The merged range that contains the cell, if any. */
  at(row: number, col: number): Range | null {
    if (!this.list.length) return null
    const bucket = this.index().get(Math.floor(row / BUCKET))
    if (!bucket) return null
    for (const m of bucket) if (row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1) return m
    return null
  }

  /** Merged ranges that overlap `range`. */
  overlapping(range: Range): Range[] {
    return this.list.filter((m) => m.r0 <= range.r1 && m.r1 >= range.r0 && m.c0 <= range.c1 && m.c1 >= range.c0)
  }

  add(range: Range) {
    this.list.push({ ...range })
    this.buckets = null
  }

  /** Removes the merges that equal one of `ranges`. */
  remove(ranges: Range[]) {
    for (const r of ranges) {
      const i = this.list.findIndex((m) => m.r0 === r.r0 && m.r1 === r.r1 && m.c0 === r.c0 && m.c1 === r.c1)
      if (i >= 0) this.list.splice(i, 1)
    }
    this.buckets = null
  }

  /** Replaces the whole list (structural edits rebuild it). */
  replace(ranges: Range[]) {
    this.list.length = 0
    this.list.push(...ranges)
    this.buckets = null
  }

  private index(): Map<number, Range[]> {
    if (this.buckets) return this.buckets
    const map = new Map<number, Range[]>()
    for (const m of this.list) {
      for (let b = Math.floor(m.r0 / BUCKET); b <= Math.floor(m.r1 / BUCKET); b++) {
        let bucket = map.get(b)
        if (!bucket) map.set(b, (bucket = []))
        bucket.push(m)
      }
    }
    this.buckets = map
    return map
  }
}

/** Where each group's summary line sits (below its detail rows, right of its columns), and the bar's visibility. */
export type OutlineSettings = { summaryBelow: boolean; summaryRight: boolean; showSymbols: boolean }

export class Sheet {
  name: string
  /** The tab's colour (Format › Tab Color), as hex. */
  tabColor?: string
  /** Hidden sheets keep their contents and formulas; their tab is not shown. */
  hidden = false
  readonly cells = new CellStore()
  readonly rows = new Axis(MAX_ROWS, DEFAULT_ROW_HEIGHT)
  readonly cols = new Axis(MAX_COLS, DEFAULT_COL_WIDTH)
  readonly rowStyles = new LineStyles(MAX_ROWS)
  readonly colStyles = new LineStyles(MAX_COLS)
  readonly merges = new Merges()
  /** Conditional formatting rules, highest priority first. */
  conditional: CFEntry[] = []
  /** Data Validation rules (a cell has at most one). */
  validations: ValidationEntry[] = []
  /** Tables (Excel's ListObjects) on the sheet. */
  tables: Table[] = []
  /** Charts on the sheet, in drawing order (the last on top). */
  charts: SheetChart[] = []
  /** How the sheet prints (Page Layout, File › Print). */
  pageSetup: PageSetup = defaultPageSetup()
  /** Notes and comments by cell. */
  readonly annotations = new Annotations()
  /** Review › Protect Sheet; null when the sheet is not protected. */
  protection: SheetProtection | null = null
  /** Data › Filter: the AutoFilter range and its column filters; null when the sheet has none. */
  autoFilter: AutoFilter | null = null
  /** Data › Outline › Settings, and whether the outline bar shows (Ctrl+8). */
  outline: OutlineSettings = { summaryBelow: true, summaryRight: true, showSymbols: true }
  scenarios: ScenarioSet = emptyScenarios()

  constructor(name: string) {
    this.name = name
  }

  /** The style id that applies to a cell: its own, else its row's, else its column's. */
  styleAt(row: number, col: number): number {
    return this.cells.style(row, col) || this.rowStyles.get(row) || this.colStyles.get(col)
  }
}
