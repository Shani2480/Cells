import {
  copyTable,
  currentRegion,
  DEFAULT_TABLE_STYLE,
  guessHeader,
  nameProblem,
  nextTableId,
  nextTableName,
  numberToEditText,
  renameTableOps,
  tableAt,
  tableCellLook,
  tableOverlaps,
  slicerColumnOps,
  uniqueColumnNames,
  unstructureOps,
  type CellStyle,
  type CellWrite,
  type Op,
  type Range,
  type Sheet,
  type StyleTable,
  type Table,
  type TotalsFunction,
  fillInputOp,
  tableRows,
  tableStyleDef,
} from "@cells/core"
import type { SheetController } from "./controller"

/** What the Create Table dialog opens with. */
export type CreateTableInit = { range: Range; header: boolean; style: string }

/** The Table Design tab's check boxes (Table Style Options). */
export type TableOption = "header" | "totals" | "rowStripes" | "columnStripes" | "firstColumn" | "lastColumn" | "filter"

export const OVERLAP_TABLE = "A table can't overlap another table."
export const OVERLAP_MERGED = "A table can't overlap a range that contains merged cells."
export const CONVERT_QUESTION = "Do you want to convert the table to a normal range?"
export const RESIZE_HEADER =
  "The headers must remain in the same row and the resulting table range must overlap the original table range."
export const NAME_TAKEN = "The name already exists. Enter a unique name."

/**
 * Tables: Insert › Table (Ctrl+T), Home › Format as Table, and the Table Design tab (name, resize, style options,
 * style, Convert to Range).
 */
export class TableCommands {
  constructor(private readonly c: SheetController) {}

  /** The table the active cell is in. */
  active(): Table | null {
    const { row, col } = this.c.selection.active
    return tableAt(this.c.sheet, row, col)
  }

  private change(ops: Op[], label: string): boolean {
    if (this.c.edit) this.c.commitEdit()
    if (!ops.length) return true
    const sel = this.c.selection
    return Boolean(this.c.transact(ops, { label, before: sel, after: sel }))
  }

  private replace(table: Table, next: Table | null): Op {
    const tables = this.c.sheet.tables.map((t) => (t === table ? next : t)).filter((t): t is Table => t !== null)
    return { kind: "tables", sheet: this.c.sheetIndex, tables }
  }

  /** Typing over a header cell: the formula changes that follow the column's new name (before the write). */
  headerRenameOps(write: Op): Op[] {
    if (write.kind !== "setCells" || write.sheet !== this.c.sheetIndex) return []
    const sheet = this.c.sheet
    const out: Op[] = []
    for (const t of sheet.tables) {
      if (!t.header) continue
      const r = t.range
      const mine = write.cells.filter((w) => w.row === r.r0 && w.col >= r.c0 && w.col <= r.c1)
      if (!mine.length) continue
      const texts = t.columns.map((_, i) => headerText(sheet.cells.get(r.r0, r.c0 + i)?.value))
      for (const w of mine) texts[w.col - r.c0] = headerText(w.data?.value)
      const names = uniqueColumnNames(texts)
      const renames: { from: string; to: string }[] = []
      names.forEach((name, i) => {
        const from = t.columns[i].name
        if (name === from) return
        out.push(...renameTableOps(this.c.wb.sheets, t, t.name, { from, to: name }))
        renames.push({ from, to: name })
      })
      out.push(...slicerColumnOps(this.c.wb.sheets, t.id, renames))
    }
    return out
  }

  /**
   * What an entry does to the tables besides going into its cell (Excel's AutoCorrect for tables): typed right under
   * a table (with no total row) or right beside it, the table grows to take it in, and new rows get the calculated
   * columns' formulas; a formula typed into an empty table column fills the column and becomes its formula.
   */
  entryOps(row: number, col: number, text: string): Op[] {
    if (!text) return []
    const sheet = this.c.sheet
    const index = this.c.sheetIndex
    const styles = this.c.wb.styles
    const inside = tableAt(sheet, row, col)
    if (inside) {
      const [d0, d1] = tableRows(inside).data
      if (!text.startsWith("=") || row < d0 || row > d1) return []
      const i = col - inside.range.c0
      for (let r = d0; r <= d1; r++) if (r !== row && sheet.cells.get(r, col)?.value !== undefined) return []
      const next = copyTable(inside)
      next.columns[i] = { ...next.columns[i], formula: shiftedFormula(sheet, styles, text, row, col, d0) }
      const fill = fillInputOp(index, sheet, styles, [{ r0: d0, r1: d1, c0: col, c1: col }], text, { row, col })
      return [fill, this.replace(inside, next)]
    }
    for (const t of sheet.tables) {
      const r = t.range
      const d1 = tableRows(t).data[1]
      if (!t.totals && row === r.r1 + 1 && col >= r.c0 && col <= r.c1) return this.growDown(t, col) ?? []
      if (col === r.c1 + 1 && row >= r.r0 && row <= d1) {
        const grown = { ...r, c1: col }
        if (tableOverlaps(sheet, grown, t) || sheet.merges.overlapping(grown).length) return []
        const next = copyTable(t)
        next.range = grown
        const typed = t.header && row === r.r0 ? headerText(text.startsWith("'") ? text.slice(1) : text) : ""
        const names = uniqueColumnNames([...t.columns.map((c) => c.name), typed])
        next.columns = [...t.columns, { name: names[names.length - 1] }]
        if (next.autoFilter) next.autoFilter = { ...next.autoFilter, range: dataWithHeader(next) }
        const ops: Op[] = [this.replace(t, next)]
        if (t.header && row !== r.r0)
          ops.push({
            kind: "setCells",
            sheet: index,
            cells: [
              {
                row: r.r0,
                col,
                data: { ...styleOf(sheet.cells.get(r.r0, col)?.style), value: names[names.length - 1] },
              },
            ],
          })
        return ops
      }
    }
    return []
  }

  /**
   * Tab (or Shift+Tab) at the edge of a table's data row: the cell it wraps to. Tab on the last cell of a table with
   * no total row adds a row first.
   */
  tabTarget(dir: "left" | "right"): { row: number; col: number } | null {
    const { row, col } = this.c.selection.active
    const t = tableAt(this.c.sheet, row, col)
    if (!t) return null
    const r = t.range
    const [d0, d1] = tableRows(t).data
    if (row < d0 || row > d1) return null
    if (dir === "left") return col === r.c0 && row > d0 ? { row: row - 1, col: r.c1 } : null
    if (col !== r.c1) return null
    if (row < d1) return { row: row + 1, col: r.c0 }
    if (t.totals) return null
    const grow = this.growDown(t)
    if (!grow) return null
    this.change(grow, "Insert Table Row")
    return { row: row + 1, col: r.c0 }
  }

  /** A table one row longer (its calculated columns filled in the new row); null when it would hit another table. */
  private growDown(t: Table, except = -1): Op[] | null {
    const sheet = this.c.sheet
    const index = this.c.sheetIndex
    const styles = this.c.wb.styles
    const r = t.range
    const row = r.r1 + 1
    const d0 = tableRows(t).data[0]
    if (tableOverlaps(sheet, { ...r, r1: row }, t)) return null
    const next = copyTable(t)
    next.range = { ...r, r1: row }
    if (next.autoFilter) next.autoFilter = { ...next.autoFilter, range: dataWithHeader(next) }
    const ops: Op[] = [this.replace(t, next)]
    t.columns.forEach((c, i) => {
      if (!c.formula || r.c0 + i === except) return
      const at = { r0: row, r1: row, c0: r.c0 + i, c1: r.c0 + i }
      ops.push(fillInputOp(index, sheet, styles, [at], "=" + c.formula, { row: d0, col: r.c0 + i }))
    })
    return ops
  }

  // ── Creating ──

  /**
   * Insert › Table (Ctrl+T or Ctrl+L) or a style from Format as Table: the Create Table dialog on the data around
   * the selection. Inside a table, a style from the gallery restyles it instead.
   */
  openCreate(style = DEFAULT_TABLE_STYLE) {
    if (this.c.edit) this.c.commitEdit()
    const inside = this.active()
    if (inside) {
      if (style !== DEFAULT_TABLE_STYLE || inside.style !== style) this.setStyle(style)
      return
    }
    const sel = this.c.selection.ranges[0]
    const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
    const range = single ? currentRegion(this.c.sheet, this.c.selection.active) : { ...sel }
    const header = guessHeader(this.c.sheet, this.c.wb.styles, range)
    this.c.shell.createTable?.({ range, header, style })
  }

  /** The dialog's OK: makes the table (a header row of Column1, Column2… goes in above the data when it has none). */
  create(range: Range, header: boolean, style: string): boolean {
    const ops = this.createOps(range, header, style)
    if (typeof ops === "string") return this.refuse(ops)
    return this.change(ops, "Create Table")
  }

  /** The changes that make a table of `range` (Quick Analysis previews them), or Excel's reason it can't. */
  createOps(range: Range, header: boolean, style: string): Op[] | string {
    const sheet = this.c.sheet
    const index = this.c.sheetIndex
    if (tableOverlaps(sheet, range)) return OVERLAP_TABLE
    if (sheet.merges.overlapping(range).length) return OVERLAP_MERGED
    const ops: Op[] = []
    const width = range.c1 - range.c0 + 1
    let r = { ...range }
    const writes: CellWrite[] = []
    if (!header) {
      // Excel moves the data down one row inside the table's columns to make room for the headers.
      ops.push({ kind: "insertCells", sheet: index, range: { r0: r.r0, r1: r.r0, c0: r.c0, c1: r.c1 }, shift: "down" })
      r = { ...r, r1: r.r1 + 1 }
    }
    const texts = Array.from({ length: width }, (_, i) =>
      header ? headerText(sheet.cells.get(r.r0, r.c0 + i)?.value) : "",
    )
    const names = uniqueColumnNames(texts)
    names.forEach((name, i) => {
      const current = header ? sheet.cells.get(r.r0, r.c0 + i) : undefined
      if (!header || current?.value !== name || current.input !== undefined)
        writes.push({
          row: r.r0,
          col: r.c0 + i,
          data: { ...(current?.style ? { style: current.style } : {}), value: name },
        })
    })
    if (writes.length) ops.push({ kind: "setCells", sheet: index, cells: writes })
    if (sheet.autoFilter && overlaps(sheet.autoFilter.range, r))
      ops.push({ kind: "autoFilter", sheet: index, filter: null })
    const table: Table = {
      id: nextTableId(this.c.wb.sheets),
      name: nextTableName(
        this.c.wb.sheets,
        this.c.wb.names.all().map((n) => n.name),
      ),
      range: r,
      header: true,
      totals: false,
      columns: names.map((name) => ({ name })),
      style,
      firstColumn: false,
      lastColumn: false,
      rowStripes: true,
      columnStripes: false,
      autoFilter: { range: { ...r }, columns: {} },
    }
    ops.push({ kind: "tables", sheet: index, tables: [...sheet.tables, table] })
    return ops
  }

  /** Rows the table's filter hid shown again (its Filter Button or header row goes). */
  private showFiltered(t: Table): Op[] {
    if (!t.autoFilter) return []
    return this.c.data.filterOps({ filter: t.autoFilter, table: t }, null).filter((op) => op.kind === "axis")
  }

  private refuse(message: string): false {
    this.c.shell.alert(message)
    return false
  }

  // ── Table Design ──

  setStyle(style: string | null) {
    const t = this.active()
    if (!t) return
    this.change([this.replace(t, { ...copyTable(t), style })], "Table Style")
  }

  /** Table Style Options' check boxes. */
  option(key: TableOption): boolean {
    const t = this.active()
    if (!t) return false
    if (key === "filter") return t.autoFilter !== null
    if (key === "header") return t.header
    if (key === "totals") return t.totals
    return t[key]
  }

  setOption(key: TableOption, on: boolean) {
    const t = this.active()
    if (!t || this.option(key) === on) return
    if (key === "totals") return this.setTotals(t, on)
    if (key === "header") return this.setHeader(t, on)
    const next = copyTable(t)
    if (key === "filter") next.autoFilter = on && t.header ? { range: dataWithHeader(t), columns: {} } : null
    else next[key] = on
    this.change([...this.showFiltered(t), this.replace(t, next)], "Table Style Options")
  }

  /** Total Row: a row under the data with "Total" and the last column's SUBTOTAL (Sum for numbers, Count otherwise). */
  private setTotals(t: Table, on: boolean) {
    const index = this.c.sheetIndex
    const sheet = this.c.sheet
    const ops: Op[] = []
    const next = copyTable(t)
    const r = t.range
    if (on) {
      const below = r.r1 + 1
      let busy = false
      for (let c = r.c0; c <= r.c1; c++) if (sheet.cells.hasValue(below, c)) busy = true
      if (busy)
        ops.push({
          kind: "insertCells",
          sheet: index,
          range: { r0: below, r1: below, c0: r.c0, c1: r.c1 },
          shift: "down",
        })
      next.range = { ...r, r1: below }
      next.totals = true
      const last = next.columns.length - 1
      const numeric = columnIsNumeric(sheet, t, last)
      next.columns = next.columns.map((c, i) => {
        if (i === 0 && last > 0) return { ...c, totals: undefined, totalsLabel: c.totalsLabel ?? "Total" }
        if (i === last) return { ...c, totals: c.totals && c.totals !== "none" ? c.totals : numeric ? "sum" : "count" }
        return c
      })
      const writes: CellWrite[] = next.columns.map((c, i) => {
        const style = sheet.cells.get(below, r.c0 + i)?.style
        return { row: below, col: r.c0 + i, data: totalsCell(c, style) }
      })
      ops.push({ kind: "tables", sheet: index, tables: sheet.tables.map((x) => (x === t ? next : x)) })
      ops.push({ kind: "setCells", sheet: index, cells: writes })
    } else {
      const writes: CellWrite[] = []
      for (let c = r.c0; c <= r.c1; c++) {
        const style = sheet.cells.get(r.r1, c)?.style
        writes.push({ row: r.r1, col: c, data: style ? { style } : undefined })
      }
      ops.push({ kind: "setCells", sheet: index, cells: writes })
      next.range = { ...r, r1: r.r1 - 1 }
      next.totals = false
      ops.push({ kind: "tables", sheet: index, tables: sheet.tables.map((x) => (x === t ? next : x)) })
    }
    if (next.autoFilter) next.autoFilter = { ...next.autoFilter, range: dataWithHeader(next) }
    this.change(ops, "Total Row")
  }

  /** Header Row: off clears the header cells and leaves them out of the table; on puts the names back above. */
  private setHeader(t: Table, on: boolean) {
    const index = this.c.sheetIndex
    const sheet = this.c.sheet
    const ops: Op[] = []
    const next = copyTable(t)
    const r = t.range
    const writes: CellWrite[] = []
    if (!on) {
      for (let c = r.c0; c <= r.c1; c++) {
        const style = sheet.cells.get(r.r0, c)?.style
        writes.push({ row: r.r0, col: c, data: style ? { style } : undefined })
      }
      ops.push({ kind: "setCells", sheet: index, cells: writes })
      next.range = { ...r, r0: r.r0 + 1 }
      next.header = false
      next.autoFilter = null
      ops.push(...this.showFiltered(t))
    } else {
      let row = r.r0 - 1
      let free = row >= 0
      for (let c = r.c0; c <= r.c1 && free; c++) if (sheet.cells.hasValue(row, c)) free = false
      if (!free) {
        ops.push({
          kind: "insertCells",
          sheet: index,
          range: { r0: r.r0, r1: r.r0, c0: r.c0, c1: r.c1 },
          shift: "down",
        })
        row = r.r0
        next.range = { ...r, r1: r.r1 + 1 }
      } else next.range = { ...r, r0: row }
      next.header = true
      next.columns.forEach((col, i) =>
        writes.push({
          row,
          col: r.c0 + i,
          data: { ...styleOf(sheet.cells.get(row, r.c0 + i)?.style), value: col.name },
        }),
      )
      next.autoFilter = { range: dataWithHeader(next), columns: {} }
    }
    // The table's new place first: the header cells written after it rename nothing.
    ops.push({ kind: "tables", sheet: index, tables: sheet.tables.map((x) => (x === t ? next : x)) })
    if (on) ops.push({ kind: "setCells", sheet: index, cells: writes })
    this.change(ops, "Header Row")
  }

  // ── Total row ──

  /** The total row's dropdown on one of its cells: Excel's functions, and the one the column uses. */
  totalsList(row: number, col: number): { items: string[]; current: string } | null {
    const t = tableAt(this.c.sheet, row, col)
    if (!t?.totals || row !== t.range.r1) return null
    const c = t.columns[col - t.range.c0]
    const current = TOTAL_FUNCTIONS.find((f) => f.id === (c.totals ?? "none"))?.label ?? "None"
    return { items: TOTAL_FUNCTIONS.map((f) => f.label), current }
  }

  /** A function picked from the total row's dropdown: the cell gets its SUBTOTAL (None empties it). */
  pickTotal(row: number, col: number, label: string) {
    const t = tableAt(this.c.sheet, row, col)
    const fn = TOTAL_FUNCTIONS.find((f) => f.label === label)
    if (!t?.totals || row !== t.range.r1 || !fn) return
    if (this.c.edit) this.c.cancelEdit()
    const i = col - t.range.c0
    const next = copyTable(t)
    const column = { ...next.columns[i], totals: fn.id === "none" ? undefined : fn.id }
    delete column.totalsLabel
    next.columns[i] = column
    const style = this.c.sheet.cells.get(row, col)?.style
    const sel = this.c.selection
    this.c.transact(
      [
        this.replace(t, next),
        { kind: "setCells", sheet: this.c.sheetIndex, cells: [{ row, col, data: totalsCell(column, style) }] },
      ],
      { label: "Total Row", before: sel, after: sel },
    )
  }

  /** Table Name: Excel's rules for names, unique among tables and defined names; formulas follow. */
  rename(name: string): boolean {
    const t = this.active()
    if (!t || name === t.name) return true
    const others = [
      ...this.c.wb.names.all(),
      ...this.c.wb.sheets.flatMap((s) => s.tables.filter((x) => x !== t).map((x) => ({ name: x.name, ref: "" }))),
    ]
    const problem = nameProblem(name, others)
    if (problem) return this.refuse(problem)
    const ops = renameTableOps(this.c.wb.sheets, t, name)
    ops.push(this.replace(t, { ...copyTable(t), name }))
    return this.change(ops, "Rename Table")
  }

  /** Resize Table: the header stays on its row, and the new range must overlap the old one. */
  resize(range: Range): boolean {
    const t = this.active()
    if (!t) return false
    const r = t.range
    const headerKept = !t.header || range.r0 === r.r0
    if (!headerKept || !overlaps(range, r) || range.r1 - range.r0 < (t.header ? 1 : 0) + (t.totals ? 1 : 0))
      return this.refuse(RESIZE_HEADER)
    if (tableOverlaps(this.c.sheet, range, t)) return this.refuse(OVERLAP_TABLE)
    const sheet = this.c.sheet
    const next = copyTable(t)
    next.range = { ...range }
    const ops: Op[] = []
    // Columns: kept by position; new ones are named from their header cells (or Column5, Column6…).
    const width = range.c1 - range.c0 + 1
    const texts = Array.from({ length: width }, (_, i) => {
      const old = range.c0 + i - r.c0
      if (old >= 0 && old < t.columns.length) return t.columns[old].name
      return t.header ? headerText(sheet.cells.get(range.r0, range.c0 + i)?.value) : ""
    })
    const names = uniqueColumnNames(texts)
    next.columns = names.map((name, i) => {
      const old = range.c0 + i - r.c0
      return old >= 0 && old < t.columns.length ? { ...t.columns[old], name } : { name }
    })
    if (t.header) {
      const writes: CellWrite[] = []
      names.forEach((name, i) => {
        const cell = sheet.cells.get(range.r0, range.c0 + i)
        if (cell?.value !== name)
          writes.push({ row: range.r0, col: range.c0 + i, data: { ...styleOf(cell?.style), value: name } })
      })
      ops.push({ kind: "tables", sheet: this.c.sheetIndex, tables: sheet.tables.map((x) => (x === t ? next : x)) })
      if (writes.length) ops.push({ kind: "setCells", sheet: this.c.sheetIndex, cells: writes })
    } else ops.push({ kind: "tables", sheet: this.c.sheetIndex, tables: sheet.tables.map((x) => (x === t ? next : x)) })
    if (next.autoFilter) next.autoFilter = { range: dataWithHeader(next), columns: {} }
    return this.change(ops, "Resize Table")
  }

  /** Convert to Range: asks first; the cells keep the table's look as their own formats, formulas lose [@…]. */
  convertToRange() {
    const t = this.active()
    if (!t) return
    this.c.shell.confirm(CONVERT_QUESTION, () => {
      const ops: Op[] = [...unstructureOps(this.c.wb.sheets, t)]
      ops.push(...this.bakeLook(t))
      ops.push(this.replace(t, null))
      this.change(ops, "Convert to Range")
    })
  }

  /** Format ops that give each cell the table style's look where it has no format of its own. */
  private bakeLook(t: Table): Op[] {
    const groups = new Map<string, { patch: Partial<CellStyle>; ranges: Range[] }>()
    const sheet = this.c.sheet
    const styles = this.c.wb.styles
    for (let row = t.range.r0; row <= t.range.r1; row++)
      for (let col = t.range.c0; col <= t.range.c1; col++) {
        const look = tableCellLook(t, row, col, tableStyleDef(t.style, this.c.wb.theme))
        if (!look) continue
        const own = styles.get(sheet.styleAt(row, col))
        const patch: Partial<CellStyle> = {}
        for (const k of Object.keys(look) as (keyof typeof look)[])
          if (own[k] === undefined) (patch as Record<string, unknown>)[k] = look[k]
        if (!Object.keys(patch).length) continue
        const key = JSON.stringify(patch)
        let g = groups.get(key)
        if (!g) groups.set(key, (g = { patch, ranges: [] }))
        g.ranges.push({ r0: row, r1: row, c0: col, c1: col })
      }
    return [...groups.values()].map((g) => ({
      kind: "format" as const,
      sheet: this.c.sheetIndex,
      ranges: mergeRuns(g.ranges),
      patch: g.patch,
    }))
  }
}

/** A column's formula as Excel keeps it: written for the first data row, without "=". */
function shiftedFormula(
  sheet: Sheet,
  styles: StyleTable,
  text: string,
  row: number,
  col: number,
  first: number,
): string {
  if (row === first) return text.slice(1)
  const op = fillInputOp(0, sheet, styles, [{ r0: first, r1: first, c0: col, c1: col }], text, { row, col })
  const input = op.kind === "setCells" ? op.cells[0]?.data?.input : undefined
  return (input ?? text).slice(1)
}

function headerText(value: unknown): string {
  if (value === undefined || value === null) return ""
  if (typeof value === "number") return numberToEditText(value)
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE"
  if (typeof value === "object" && "error" in (value as object)) return (value as { error: string }).error
  return String(value)
}

const styleOf = (style: number | undefined) => (style ? { style } : {})
const overlaps = (a: Range, b: Range) => a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/** The filter's range: the header and data rows. */
const dataWithHeader = (t: Table): Range => ({ ...t.range, r1: t.range.r1 - (t.totals ? 1 : 0) })

/** Whether a column's data is all numbers (the total row then sums it). */
function columnIsNumeric(sheet: SheetController["sheet"], t: Table, i: number): boolean {
  const col = t.range.c0 + i
  const r0 = t.range.r0 + (t.header ? 1 : 0)
  const r1 = t.range.r1 - (t.totals ? 1 : 0)
  let any = false
  for (let r = r0; r <= r1; r++) {
    const v = sheet.cells.get(r, col)?.value
    if (v === undefined || v === "") continue
    if (typeof v !== "number") return false
    any = true
  }
  return any
}

/** The total row dropdown's functions, in Excel's order ("More Functions..." is left out). */
export const TOTAL_FUNCTIONS: { id: TotalsFunction; label: string }[] = [
  { id: "none", label: "None" },
  { id: "average", label: "Average" },
  { id: "count", label: "Count" },
  { id: "countNums", label: "Count Numbers" },
  { id: "max", label: "Max" },
  { id: "min", label: "Min" },
  { id: "sum", label: "Sum" },
  { id: "stdDev", label: "StdDev" },
  { id: "var", label: "Var" },
]

/** The SUBTOTAL codes Excel's total row uses (101–111 leave filtered rows out). */
const SUBTOTAL: Record<string, number> = {
  average: 101,
  count: 103,
  countNums: 102,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
}

/** A total row cell for a column: its label, its SUBTOTAL, or nothing. */
export function totalsCell(c: Table["columns"][number], style?: number) {
  const base = style ? { style } : {}
  if (c.totals && SUBTOTAL[c.totals]) {
    const ref = /^[A-Za-z0-9_]+$/.test(c.name) ? `[${c.name}]` : `[[${c.name.replace(/['[\]#]/g, "'$&")}]]`
    return { ...base, value: 0, input: `=SUBTOTAL(${SUBTOTAL[c.totals]},${ref})` }
  }
  if (c.totalsLabel) return { ...base, value: c.totalsLabel }
  return style ? base : undefined
}

/** Single cells joined into row runs (fewer ranges for a format op). */
function mergeRuns(cells: Range[]): Range[] {
  const out: Range[] = []
  for (const r of cells) {
    const last = out[out.length - 1]
    if (last && last.r0 === r.r0 && last.r1 === r.r1 && last.c1 + 1 === r.c0) last.c1 = r.c1
    else out.push({ ...r })
  }
  return out
}
