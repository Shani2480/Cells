import {
  cellName,
  contentForInput,
  displayText,
  findAll,
  findRegExp,
  replacedText,
  searchAreas,
  searchText,
  selectCell,
  specialCells,
  type CellAt,
  type CellWrite,
  type FindOptions,
  type Op,
  type Selection,
  type SpecialKind,
} from "@cells/core"
import type { SheetController } from "./controller"

/** Find and Replace's options: Find's plus Within (the sheet or every sheet). */
export type SearchOptions = FindOptions & { within: "sheet" | "workbook" }

/** One row of Find All's list. */
export type FoundCell = {
  sheet: number
  sheetName: string
  row: number
  col: number
  cell: string
  value: string
  formula: string
}

export const NOT_FOUND = "We couldn't find what you were looking for. Select Options for more ways to search."
export const NOTHING_TO_REPLACE = "We couldn't find anything to replace. Select Options for more ways to search."

/**
 * Home › Find & Select: Find (Ctrl+F), Replace (Ctrl+H), Find Next (Shift+F4) and Go To Special, as Excel does them.
 * Searching looks in the selection when it is more than one cell, else the whole sheet (or every sheet).
 */
export class FindCommands {
  /** The last search, for Shift+F4 and for the dialog to open with. */
  last: SearchOptions = { text: "", lookIn: "formulas", within: "sheet" }
  replaceWith = ""

  constructor(private readonly c: SheetController) {}

  private sheetOrder(o: SearchOptions): number[] {
    const n = this.c.wb.sheets.length
    if (o.within === "sheet") return [this.c.sheetIndex]
    return Array.from({ length: n }, (_, i) => (this.c.sheetIndex + i) % n).filter((i) => !this.c.wb.sheets[i].hidden)
  }

  /** The areas searched on a sheet: the selection on the shown sheet when it is more than one cell. */
  private areasOn(index: number) {
    const sheet = this.c.wb.sheets[index]
    const single = this.c.selection.ranges.length === 1 && isOneCell(this.c.selection.ranges[0])
    if (index === this.c.sheetIndex && !single) return searchAreas(sheet, this.c.selection.ranges)
    return searchAreas(sheet, [{ r0: 0, c0: 0, r1: 0, c1: 0 }])
  }

  /** Puts the active cell on a found cell: inside the selection when searching it, else selecting just the cell. */
  private show(index: number, at: CellAt) {
    if (index !== this.c.sheetIndex) this.c.activate(index)
    const sel = this.c.selection
    const inside = sel.ranges.findIndex(
      (r) => !isOneCell(r) && at.row >= r.r0 && at.row <= r.r1 && at.col >= r.c0 && at.col <= r.c1,
    )
    const next: Selection =
      inside >= 0 ? { ...sel, active: at, activeRange: inside } : selectCell({ row: at.row, col: at.col })
    this.c.setSelection(next, at)
  }

  /** Find Next (or previous); false (with Excel's message) when nothing matches. */
  next(o: SearchOptions = this.last, backwards = false): boolean {
    this.last = o
    if (this.c.edit) this.c.commitEdit()
    if (!o.text) return false
    const order = this.sheetOrder(o)
    const styles = this.c.wb.styles
    const from = this.c.selection.active
    const pick = (list: CellAt[]) => (backwards ? list[list.length - 1] : list[0])
    // The shown sheet after the active cell, then the other sheets in tab order, then round to the start.
    const here = findAll(this.c.sheet, styles, o, this.areasOn(this.c.sheetIndex))
    const after = backwards ? here.filter((c) => cmp(o, c, from) < 0) : here.filter((c) => cmp(o, c, from) > 0)
    if (after.length) {
      this.show(this.c.sheetIndex, pick(after))
      return true
    }
    for (const index of backwards ? order.slice(1).reverse() : order.slice(1)) {
      const list = findAll(this.c.wb.sheets[index], styles, o, this.areasOn(index))
      if (list.length) {
        this.show(index, pick(list))
        return true
      }
    }
    if (here.length) {
      this.show(this.c.sheetIndex, pick(here))
      return true
    }
    this.c.shell.alert(NOT_FOUND)
    return false
  }

  /** Find All: every match, sheet by sheet. */
  all(o: SearchOptions = this.last): FoundCell[] {
    this.last = o
    if (!o.text) return []
    const out: FoundCell[] = []
    for (const index of this.sheetOrder(o).sort((a, b) => a - b)) {
      const sheet = this.c.wb.sheets[index]
      for (const at of findAll(sheet, this.c.wb.styles, o, this.areasOn(index))) {
        const data = sheet.cells.get(at.row, at.col)
        const code = this.c.wb.styles.get(sheet.styleAt(at.row, at.col)).numberFormat
        out.push({
          sheet: index,
          sheetName: sheet.name,
          row: at.row,
          col: at.col,
          cell: "$" + cellName(at).replace(/(\d+)$/, "$$$1"),
          value: displayText(data?.value, code),
          formula: data?.input?.startsWith("=") ? data.input : "",
        })
      }
    }
    return out
  }

  /** Shows a cell from Find All's list. */
  goToFound(found: FoundCell) {
    this.show(found.sheet, found)
  }

  /** The writes that replace matches in `cells` of a sheet (typed again, so numbers stay numbers). */
  private writes(index: number, cells: CellAt[], o: SearchOptions, replacement: string): CellWrite[] {
    const sheet = this.c.wb.sheets[index]
    const styles = this.c.wb.styles
    const out: CellWrite[] = []
    for (const at of cells) {
      const text = searchText(sheet, styles, at.row, at.col, "formulas")
      if (text === null) continue
      const next = replacedText(text, o, replacement)
      if (next === null || next === text) continue
      out.push({ row: at.row, col: at.col, data: contentForInput(styles, sheet.cells.get(at.row, at.col), next) })
    }
    return out
  }

  /** Replace: the active cell when it matches, then on to the next match. */
  replace(o: SearchOptions, replacement: string): boolean {
    this.last = { ...o, lookIn: "formulas" }
    this.replaceWith = replacement
    if (this.c.edit) this.c.commitEdit()
    const opts = this.last
    const { row, col } = this.c.selection.active
    const text = searchText(this.c.sheet, this.c.wb.styles, row, col, "formulas")
    if (text !== null && findRegExp(opts).test(text)) {
      const cells = this.writes(this.c.sheetIndex, [{ row, col }], opts, replacement)
      if (cells.length) {
        const before = this.c.selection
        this.c.transact([{ kind: "setCells", sheet: this.c.sheetIndex, cells }], {
          label: "Replace",
          before,
          after: before,
        })
      }
      return this.next(opts)
    }
    if (!this.next(opts)) return false
    return true
  }

  /** Replace All: every match in one undo step; returns how many cells changed (Excel's "We made N replacements"). */
  replaceAll(o: SearchOptions, replacement: string): number {
    this.last = { ...o, lookIn: "formulas" }
    this.replaceWith = replacement
    if (this.c.edit) this.c.commitEdit()
    const ops: Op[] = []
    let count = 0
    for (const index of this.sheetOrder(this.last)) {
      const cells = this.writes(
        index,
        findAll(this.c.wb.sheets[index], this.c.wb.styles, this.last, this.areasOn(index)),
        this.last,
        replacement,
      )
      count += cells.length
      if (cells.length) ops.push({ kind: "setCells", sheet: index, cells })
    }
    if (!ops.length) {
      this.c.shell.alert(NOTHING_TO_REPLACE)
      return 0
    }
    const before = this.c.selection
    if (!this.c.transact(ops, { label: "Replace All", before, after: before })) return 0
    this.c.shell.alert(`All done. We made ${count} replacement${count === 1 ? "" : "s"}.`)
    return count
  }

  /** Go To Special: selects the cells; "No cells were found." when there are none. */
  special(kind: SpecialKind): boolean {
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    const ranges = specialCells(this.c.sheet, this.c.wb.styles, kind, sel.ranges, sel.active)
    if (!ranges.length) {
      this.c.shell.alert("No cells were found.")
      return false
    }
    // Excel keeps the active cell when it is among the cells picked (Current Region, Visible Cells).
    const keep = ranges.findIndex(
      (r) => sel.active.row >= r.r0 && sel.active.row <= r.r1 && sel.active.col >= r.c0 && sel.active.col <= r.c1,
    )
    const activeRange = keep >= 0 ? keep : 0
    const active = keep >= 0 ? sel.active : { row: ranges[0].r0, col: ranges[0].c0 }
    this.c.setSelection({ ranges, active, activeRange, anchor: active, focus: active }, active)
    return true
  }
}

const isOneCell = (r: { r0: number; r1: number; c0: number; c1: number }) => r.r0 === r.r1 && r.c0 === r.c1
const cmp = (o: FindOptions, a: CellAt, b: CellAt) =>
  o.byColumns ? a.col - b.col || a.row - b.row : a.row - b.row || a.col - b.col
