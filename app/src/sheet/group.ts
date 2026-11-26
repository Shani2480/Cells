import {
  MAX_COLS,
  MAX_ROWS,
  rangeIntersection,
  type AxisEntry,
  type CellData,
  type CellWrite,
  type Op,
  type Range,
  type Sheet,
  type StyleRecord,
} from "@cells/core"
import type { SheetController } from "./controller"

/**
 * Grouped sheets (Excel's "[Group]"): Shift+click or Ctrl+click on sheet tabs, Select All Sheets, Ctrl+Shift+PgDn/PgUp.
 * While sheets are grouped, what is typed, pasted, cleared or formatted on the active sheet happens on each grouped
 * sheet too, in the same cells, as one undo step. Features that belong to one sheet (tables, charts, filters,
 * conditional formatting, validation, notes and comments, links) are refused, as Excel greys them out.
 */

/** Edits that repeat on every grouped sheet. */
const REPEATED: ReadonlySet<Op["kind"]> = new Set([
  "setCells",
  "format",
  "styles",
  "merge",
  "unmerge",
  "axis",
  "insertRows",
  "deleteRows",
  "insertCols",
  "deleteCols",
  "insertCells",
  "deleteCells",
  "pageSetup",
  "outlineLevels",
])

/** Changes that work on the whole workbook or on sheets as a whole: they go through as they are. */
const WORKBOOK: ReadonlySet<Op["kind"]> = new Set([
  "addSheet",
  "removeSheet",
  "renameSheet",
  "moveSheet",
  "sheetProps",
  "names",
  "externalLinks",
  "protectWorkbook",
  "theme",
])

/** What Cells says when a one-sheet feature is used on a group (Excel greys these commands out). */
export const GROUP_MESSAGE =
  "This command can't be used while several sheets are grouped. Click a single sheet tab (or right-click a tab and choose Ungroup Sheets) and try again."

export class SheetGroup {
  /** The grouped sheets besides the active one; empty: one sheet selected. */
  private readonly others = new Set<Sheet>()

  constructor(private readonly c: SheetController) {}

  /** Indexes of the grouped sheets, the active one among them, in tab order. */
  indexes(): number[] {
    const out: number[] = []
    this.c.wb.sheets.forEach((s, i) => (i === this.c.sheetIndex || this.others.has(s)) && out.push(i))
    return out
  }

  get active(): boolean {
    return this.indexes().length > 1
  }

  has(index: number): boolean {
    const sheet = this.c.wb.sheets[index]
    return index === this.c.sheetIndex || (sheet !== undefined && this.others.has(sheet))
  }

  /** Whether every visible sheet is in the group (Select All Sheets). */
  get all(): boolean {
    return this.c.wb.sheets.every((s, i) => s.hidden || this.has(i))
  }

  /** Makes these sheets the group (the active sheet stays active and is always in it). */
  set(indexes: number[]) {
    this.others.clear()
    for (const i of indexes) {
      const sheet = this.c.wb.sheets[i]
      if (sheet && !sheet.hidden && i !== this.c.sheetIndex) this.others.add(sheet)
    }
    this.c.changed()
  }

  /** Forgets the group (a sheet shown that isn't in it, Ungroup Sheets). */
  clear(notify = true) {
    if (!this.others.size) return
    this.others.clear()
    if (notify) this.c.changed()
  }

  /** The active sheet moves to `index`, which stays in the group (the old active sheet does too). */
  keepWhileActivating(index: number) {
    const from = this.c.wb.sheets[this.c.sheetIndex]
    const to = this.c.wb.sheets[index]
    if (!from || !to || !this.others.has(to)) return
    this.others.delete(to)
    this.others.add(from)
  }

  /** Whether showing `index` keeps the group: it is in the group (Excel ungroups on a tab outside it). */
  keeps(index: number): boolean {
    const sheet = this.c.wb.sheets[index]
    return sheet !== undefined && this.others.has(sheet)
  }

  /**
   * A click on a tab. Plain: shows the sheet (in a partial group the group stays; when every sheet is grouped, a click
   * on another tab ungroups, as Excel). Ctrl: adds the sheet to the group or takes it out. Shift: groups the sheets
   * from the active one to this one.
   */
  click(index: number, mod: { ctrl: boolean; shift: boolean }) {
    const c = this.c
    if (mod.shift) {
      const lo = Math.min(index, c.sheetIndex)
      const hi = Math.max(index, c.sheetIndex)
      const range: number[] = []
      for (let i = lo; i <= hi; i++) if (!c.wb.sheets[i]?.hidden) range.push(i)
      this.set(range)
      return
    }
    if (mod.ctrl) {
      if (index === c.sheetIndex) return
      const sheet = c.wb.sheets[index]
      if (!sheet) return
      if (this.others.has(sheet)) this.others.delete(sheet)
      else this.others.add(sheet)
      c.changed()
      return
    }
    if (this.all && this.active && index !== c.sheetIndex) this.clear(false)
    c.activate(index)
  }

  /** Select All Sheets (the tab menu). */
  selectAll() {
    this.set(this.c.sheets.visible())
  }

  /** Ctrl+Shift+PgDn / PgUp: the next or previous visible sheet joins the group and becomes active (Excel). */
  extend(dir: 1 | -1) {
    const c = this.c
    const sheets = c.wb.sheets
    const keep = this.indexes()
    for (let i = c.sheetIndex + dir; i >= 0 && i < sheets.length; i += dir) {
      if (sheets[i].hidden) continue
      if (!keep.includes(i)) keep.push(i)
      const group = keep.map((k) => sheets[k])
      this.others.clear()
      c.activate(i)
      for (const s of group) if (s !== sheets[c.sheetIndex]) this.others.add(s)
      c.changed()
      return
    }
  }

  /**
   * The ops a change makes on the group: edits on the active sheet repeat on each grouped sheet. Null: refused (a
   * one-sheet feature). Changes worked out per sheet (`each`) skip this.
   */
  expand(ops: Op[]): Op[] | null {
    const group = this.indexes()
    if (group.length < 2) return ops
    const at = this.c.sheetIndex
    if (ops.some((op) => !REPEATED.has(op.kind) && !WORKBOOK.has(op.kind))) return null
    const out = [...ops]
    for (const g of group) {
      if (g === at) continue
      for (const op of ops)
        if (REPEATED.has(op.kind) && "sheet" in op && op.sheet === at) out.push({ ...op, sheet: g } as Op)
    }
    return out
  }

  /** One op per grouped sheet, each worked out from that sheet's own cells (Delete clears what each sheet has). */
  each(make: (index: number, sheet: Sheet) => Op | Op[] | null): Op[] {
    const ops: Op[] = []
    for (const i of this.indexes()) {
      const op = make(i, this.c.wb.sheets[i])
      if (op) ops.push(...(Array.isArray(op) ? op : [op]))
    }
    return ops
  }

  /**
   * Home › Fill › Across Worksheets: the selection's cells copied from the active sheet to the same cells of each
   * grouped sheet: everything, the contents only (each cell keeps its formats) or the formats only.
   */
  fillAcross(what: "all" | "contents" | "formats") {
    const c = this.c
    if (c.edit) c.commitEdit()
    const at = c.sheetIndex
    const from = c.sheet
    const ops: Op[] = []
    for (const g of this.indexes()) {
      if (g === at) continue
      const to = c.wb.sheets[g]
      const cells: CellWrite[] = []
      for (const range of c.selection.ranges) {
        // Only where either sheet has something (a whole column copies its used part).
        const bounds = [from.cells.usedBounds(), to.cells.usedBounds()].filter((b) => b !== null)
        if (!bounds.length) continue
        const used = bounds.reduce((a, b) => ({
          r0: Math.min(a.r0, b.r0),
          c0: Math.min(a.c0, b.c0),
          r1: Math.max(a.r1, b.r1),
          c1: Math.max(a.c1, b.c1),
        }))
        const area = rangeIntersection(range, used)
        if (!area) continue
        for (let row = area.r0; row <= area.r1; row++)
          for (let col = area.c0; col <= area.c1; col++) {
            const source = from.cells.get(row, col)
            const target = to.cells.get(row, col)
            if (!source && !target) continue
            const data = filled(source, target, what)
            cells.push({ row, col, data })
          }
      }
      if (cells.length) ops.push({ kind: "setCells", sheet: g, cells })
      // Whole rows or columns take their format and size with them (Excel 365 copies both with All and Formats).
      if (what !== "contents") ops.push(...lineOps(from, to, g, c.selection.ranges))
    }
    if (!ops.length) return
    c.transact(ops, { label: "Fill Across Worksheets", before: c.selection, after: c.selection, perSheet: true })
  }
}

/** The row and column formats and sizes Fill Across Worksheets copies for whole rows and columns in `ranges`. */
function lineOps(from: Sheet, to: Sheet, sheet: number, ranges: readonly Range[]): Op[] {
  const ops: Op[] = []
  const record: StyleRecord = { cells: [], rows: [], cols: [] }
  for (const axis of ["rows", "cols"] as const) {
    const entries: AxisEntry[] = []
    const styles = axis === "rows" ? record.rows : record.cols
    const [src, dst] = axis === "rows" ? [from.rows, to.rows] : [from.cols, to.cols]
    const [srcStyles, dstStyles] = axis === "rows" ? [from.rowStyles, to.rowStyles] : [from.colStyles, to.colStyles]
    for (const r of ranges) {
      const whole = axis === "rows" ? r.c0 === 0 && r.c1 === MAX_COLS - 1 : r.r0 === 0 && r.r1 === MAX_ROWS - 1
      if (!whole) continue
      const [lo, hi] = axis === "rows" ? [r.r0, r.r1] : [r.c0, r.c1]
      for (let i = lo; i <= hi; i++) {
        const style = srcStyles.get(i)
        if (style !== dstStyles.get(i)) styles.push(i, style)
        const size = src.hasCustomSize(i) ? src.storedSize(i) : null
        if (size !== (dst.hasCustomSize(i) ? dst.storedSize(i) : null))
          entries.push({ index: i, size, hidden: dst.isHidden(i) })
      }
    }
    if (entries.length) ops.push({ kind: "axis", sheet, axis, entries })
  }
  if (record.rows.length || record.cols.length) ops.push({ kind: "styles", sheet, record })
  return ops
}

/** One cell after Fill Across Worksheets: `source` from the active sheet over `target`. */
function filled(source: CellData | undefined, target: CellData | undefined, what: "all" | "contents" | "formats") {
  const content = (d: CellData | undefined): CellData => {
    const out: CellData = {}
    if (d?.value !== undefined && !d.spill) out.value = d.value
    if (d?.input !== undefined) out.input = d.input
    return out
  }
  const style = (d: CellData | undefined) => (d?.style ? { style: d.style } : {})
  const data =
    what === "all"
      ? { ...content(source), ...style(source) }
      : what === "contents"
        ? { ...content(source), ...style(target) }
        : { ...content(target), ...style(source) }
  return Object.keys(data).length ? data : undefined
}
