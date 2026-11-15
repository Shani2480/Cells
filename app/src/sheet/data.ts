import {
  copyTable,
  displayText,
  flashFill,
  type CellWrite,
  type Table,
  type AutoFilter,
  type ColumnFilter,
  currentRegion,
  filterRowEntries,
  type Op,
  guessHeader,
  isDateFormat,
  removeDuplicatesOps,
  selectRange,
  SortError,
  regional,
  sortOps,
  type Range,
  type SortOptions,
} from "@cells/core"
import type { TextImport } from "@cells/io"
import { todaySerial, type SheetController } from "./controller"
import {
  columnTexts,
  NOTHING_TO_PARSE,
  ONE_COLUMN_ONLY,
  REPLACE_DESTINATION,
  TEXT_COLUMNS_DEFAULT,
  textColumnsRange,
  textToColumnsOps,
} from "./text-columns"

/** What the Custom AutoFilter dialog opens with: the first condition's operator (and a second for Between). */
export type CustomPreset = { first: CustomOperator; second?: CustomOperator; and?: boolean }

/** The Custom AutoFilter dialog's operators, as Excel lists them. */
export type CustomOperator =
  | "equals"
  | "notEquals"
  | "greater"
  | "greaterOrEqual"
  | "less"
  | "lessOrEqual"
  | "beginsWith"
  | "notBeginsWith"
  | "endsWith"
  | "notEndsWith"
  | "contains"
  | "notContains"

export type FilterDialogRequest =
  { kind: "custom"; col: number; preset?: CustomPreset } | { kind: "top10"; col: number }

/** What a sort works on: the range, whether its first row is headers, and the column the active cell is in. */
export type SortTarget = { range: Range; header: boolean }

/** Excel's names for a quick sort, by what the active cell's column holds. */
export type SortLabels = { ascending: string; descending: string }

/** Excel's message when Flash Fill finds no pattern. */
export const NO_FLASH_PATTERN =
  "We looked at all the data next to the selection and didn't see a pattern for filling in values for you. To use Flash Fill, enter a couple of examples of the output you'd like to see, keep the active cell in the column you want filled in, and click the Flash Fill button again."

/** A filter and where it lives: the sheet's AutoFilter (table null) or a table's Filter Button. */
export type FilterHost = { filter: AutoFilter; table: Table | null }

/**
 * Data › Sort & Filter: Sort A to Z / Z to A on the data around the active cell, the Sort dialog's sort, and the
 * AutoFilter (its dropdowns, Clear and Reapply).
 */
export class DataCommands {
  constructor(private readonly c: SheetController) {}

  /** The current region for one cell; the selection itself for more than one. */
  sortRange(): Range {
    const { selection } = this.c
    const range = selection.ranges[0]
    if (range.r0 === range.r1 && range.c0 === range.c1) return currentRegion(this.c.sheet, selection.active)
    return range
  }

  /** Whether data continues past the selection (Excel then asks whether to expand it). */
  dataBeside(range: Range): boolean {
    const region = currentRegion(this.c.sheet, { row: range.r0, col: range.c0 })
    return region.c0 < range.c0 || region.c1 > range.c1
  }

  target(range = this.sortRange()): SortTarget {
    return { range, header: guessHeader(this.c.sheet, this.c.wb.styles, range) }
  }

  /** "Sort A to Z", "Sort Smallest to Largest" or "Sort Oldest to Newest", for the active cell. */
  labels(): SortLabels {
    const { row, col } = this.c.selection.active
    const data = this.c.sheet.cells.get(row, col)
    const code = this.c.wb.styles.get(this.c.sheet.styleAt(row, col)).numberFormat
    if (typeof data?.value === "number")
      return code && isDateFormat(code)
        ? { ascending: "Sort Oldest to Newest", descending: "Sort Newest to Oldest" }
        : { ascending: "Sort Smallest to Largest", descending: "Sort Largest to Smallest" }
    return { ascending: "Sort A to Z", descending: "Sort Z to A" }
  }

  /** Sort A to Z (Z to A) on the active cell's column. Asks first when the selection has data beside it. */
  quickSort(descending: boolean) {
    const selected = this.sortRange()
    const single = this.c.selection.ranges[0]
    const multi = single.r0 !== single.r1 || single.c0 !== single.c1
    const go = (range: Range) => {
      const t = this.target(range)
      this.sort(t.range, { header: t.header, keys: [{ index: this.c.selection.active.col, descending }] })
    }
    if (multi && this.dataBeside(selected)) {
      this.c.shell.sortWarning?.((expand) =>
        go(expand ? currentRegion(this.c.sheet, this.c.selection.active) : selected),
      )
      return
    }
    go(selected)
  }

  /** Data › Sort: the Sort dialog on the data around the selection (asking first when there's more beside it). */
  openSortDialog() {
    const selected = this.sortRange()
    const single = this.c.selection.ranges[0]
    const multi = single.r0 !== single.r1 || single.c0 !== single.c1
    if (multi && this.dataBeside(selected)) {
      this.c.shell.sortWarning?.((expand) =>
        this.c.shell.sortDialog?.(expand ? currentRegion(this.c.sheet, this.c.selection.active) : selected),
      )
      return
    }
    this.c.shell.sortDialog?.(selected)
  }

  /** Sorts `range`, keeping the selection on it (the whole range selected afterwards, as Excel does). */
  sort(range: Range, options: SortOptions): boolean {
    if (this.c.edit) this.c.commitEdit()
    let ops
    try {
      ops = sortOps(this.c.sheetIndex, this.c.sheet, this.c.wb.styles, range, options)
    } catch (error) {
      if (error instanceof SortError) {
        this.c.shell.alert(error.message)
        return false
      }
      throw error
    }
    const after = selectRange(range, this.c.selection.active)
    if (!ops.length) {
      this.c.setSelection(after, null)
      return true
    }
    const entry = this.c.transact(ops, { label: "Sort", before: this.c.selection, after })
    if (!entry) return false
    this.c.setSelection(after, null)
    return true
  }

  // ── Remove Duplicates ──

  /** Data › Remove Duplicates: the dialog on the data around the selection (asking first when there's more beside it). */
  openRemoveDuplicates() {
    const selected = this.sortRange()
    const single = this.c.selection.ranges[0]
    const multi = single.r0 !== single.r1 || single.c0 !== single.c1
    const open = (range: Range) => {
      this.c.setSelection(selectRange(range, this.c.selection.active), null)
      this.c.shell.removeDuplicates?.(range)
    }
    if (multi && this.dataBeside(selected)) {
      this.c.shell.sortWarning?.(
        (expand) => open(expand ? currentRegion(this.c.sheet, this.c.selection.active) : selected),
        "duplicates",
      )
      return
    }
    open(selected)
  }

  /** Removes the rows that repeat the chosen columns, in one undo step, and says how many went (as Excel). */
  removeDuplicates(range: Range, columns: number[], header: boolean): number {
    if (this.c.edit) this.c.commitEdit()
    const result = removeDuplicatesOps(this.c.sheetIndex, this.c.sheet, this.c.wb.styles, range, columns, header)
    if (result.ops.length) {
      const sel = this.c.selection
      if (!this.c.transact(result.ops, { label: "Remove Duplicates", before: sel, after: sel })) return 0
    }
    this.c.shell.alert(
      result.removed
        ? `${result.removed} duplicate value${result.removed === 1 ? "" : "s"} found and removed; ${result.remain} unique value${result.remain === 1 ? "" : "s"} remain${result.remain === 1 ? "s" : ""}.`
        : "No duplicate values found.",
    )
    return result.removed
  }

  // ── Text to Columns ──

  /** The wizard's last choices, which it opens with next time (Excel keeps them for the session). */
  /** The wizard's last choices this session (Excel keeps them); null: the defaults, with the region's separators. */
  private lastTextColumns: TextImport | null = null
  get textColumns(): TextImport {
    const r = regional()
    return this.lastTextColumns ?? { ...TEXT_COLUMNS_DEFAULT, decimal: r.decimal, thousands: r.group || " " }
  }
  set textColumns(choices: TextImport) {
    this.lastTextColumns = choices
  }

  /** Data › Text to Columns: one column of cells (the active cell's column of its region), then the wizard. */
  openTextToColumns() {
    if (this.c.edit) this.c.commitEdit()
    const { ranges, active } = this.c.selection
    if (ranges.length > 1 || ranges[0].c0 !== ranges[0].c1) {
      this.c.shell.alert(ONE_COLUMN_ONLY)
      return
    }
    const range = textColumnsRange(this.c.sheet, ranges[0], active)
    const texts = columnTexts(this.c.sheet, this.c.wb.styles, range)
    if (texts.every((t) => t === "")) {
      this.c.shell.alert(NOTHING_TO_PARSE)
      return
    }
    this.c.setSelection(selectRange(range, active), null)
    this.c.shell.textToColumns?.(range, texts)
  }

  /** Splits the column into the cells from `dest`, in one undo step; asks first when that replaces other data. */
  textToColumns(range: Range, choices: TextImport, dest: { row: number; col: number }) {
    this.textColumns = { ...choices, columns: [] }
    const { ops, overwrites } = textToColumnsOps(
      this.c.sheetIndex,
      this.c.sheet,
      this.c.wb.styles,
      range,
      choices,
      dest,
    )
    if (!ops.length) return
    const run = () => {
      const sel = this.c.selection
      this.c.transact(ops, { label: "Text to Columns", before: sel, after: sel })
    }
    if (overwrites) this.c.shell.confirm(REPLACE_DESTINATION, run)
    else run()
  }

  // ── Flash Fill ──

  /**
   * Data › Flash Fill (Ctrl+E): fills the empty cells of the active cell's column, beside the data, by the pattern of
   * the cells already typed there (its examples) from the other columns of the same rows. The results are text, as
   * Excel's. Returns how many cells were filled.
   */
  flashFill(): number {
    if (this.c.edit) this.c.commitEdit()
    const sheet = this.c.sheet
    const styles = this.c.wb.styles
    const { row, col } = this.c.selection.active
    const region = currentRegion(sheet, { row, col })
    const text = (r: number, c: number) => {
      const v = sheet.cells.get(r, c)?.value
      return v === undefined || v === "" ? null : displayText(v, styles.get(sheet.styleAt(r, c)).numberFormat)
    }
    const cols: number[] = []
    for (let c = region.c0; c <= region.c1; c++) if (c !== col) cols.push(c)
    const attempt = (r0: number): { rows: number[]; results: (string | null)[] } | null => {
      const rows: number[] = []
      for (let r = r0; r <= region.r1; r++) rows.push(r)
      const sources = rows.map((r) => cols.map((c) => text(r, c) ?? ""))
      const outputs = rows.map((r) => text(r, col))
      if (!outputs.some((o) => o !== null) || !outputs.some((o) => o === null)) return null
      const results = flashFill(sources, outputs)
      return results ? { rows, results } : null
    }
    // A header row over the column (First Name) isn't an example: without it when it doesn't fit.
    const found = cols.length ? (attempt(region.r0) ?? attempt(region.r0 + 1)) : null
    const writes: CellWrite[] = []
    found?.rows.forEach((r, i) => {
      const value = found.results[i]
      if (value === null || value === undefined) return
      const style = sheet.cells.get(r, col)?.style
      writes.push({ row: r, col, data: { ...(style ? { style } : {}), value } })
    })
    if (!writes.length) {
      this.c.shell.alert(NO_FLASH_PATTERN)
      return 0
    }
    const sel = this.c.selection
    this.c.transact([{ kind: "setCells", sheet: this.c.sheetIndex, cells: writes }], {
      label: "Flash Fill",
      before: sel,
      after: sel,
    })
    return writes.length
  }

  // ── AutoFilter ──

  /** Set by the grid: opens the dropdown of a filter header cell (Alt+Down on it). */
  openMenu: ((row: number, col: number) => void) | null = null
  /** The header cell whose dropdown (or a dialog from it) is open: commands work on that filter. */
  focus: { row: number; col: number } | null = null

  /** Every filter on the sheet: the sheet's AutoFilter and each table's Filter Button. */
  hosts(): FilterHost[] {
    const sheet = this.c.sheet
    const out: FilterHost[] = sheet.autoFilter ? [{ filter: sheet.autoFilter, table: null }] : []
    for (const table of sheet.tables) if (table.autoFilter) out.push({ filter: table.autoFilter, table })
    return out
  }

  /** The filter whose header row has this cell. */
  hostAt(row: number, col: number): FilterHost | null {
    return (
      this.hosts().find((h) => h.filter.range.r0 === row && col >= h.filter.range.c0 && col <= h.filter.range.c1) ??
      null
    )
  }

  /**
   * The filter commands work on: the one whose dropdown is open, else the active cell's table, else the sheet's
   * AutoFilter (Data › Clear and Reapply, the status bar's count).
   */
  host(): FilterHost | null {
    if (this.focus) {
      const h = this.hostAt(this.focus.row, this.focus.col)
      if (h) return h
    }
    const t = this.c.tables.active()
    if (t?.autoFilter) return { filter: t.autoFilter, table: t }
    return this.c.sheet.autoFilter ? { filter: this.c.sheet.autoFilter, table: null } : null
  }

  /** Alt+Down on a header of a filter: its dropdown. False when the active cell isn't one. */
  openMenuAtActive(): boolean {
    const { row, col } = this.c.selection.active
    if (!this.openMenu || !this.hostAt(row, col)) return false
    this.openMenu(row, col)
    return true
  }

  get autoFilter(): AutoFilter | null {
    return this.host()?.filter ?? null
  }

  /**
   * The range Filter turns on for: the selection, or the data around the active cell when one cell is selected. A
   * selected row of headers takes the data under it. Null on an empty cell with nothing around it.
   */
  filterRange(): Range | null {
    const sheet = this.c.sheet
    const sel = this.c.selection.ranges[0]
    if (sel.r0 === sel.r1 && sel.c0 === sel.c1) {
      const region = currentRegion(sheet, this.c.selection.active)
      const lone = region.r0 === region.r1 && region.c0 === region.c1
      if (lone && sheet.cells.get(region.r0, region.c0)?.value === undefined) return null
      return region
    }
    if (sel.r0 === sel.r1) {
      const region = currentRegion(sheet, { row: sel.r0, col: sel.c0 })
      return { ...sel, r1: Math.max(sel.r1, region.r1) }
    }
    return sel
  }

  /** Ctrl+Shift+L: turns the AutoFilter on for the data, or off (showing every row it hid). */
  toggleFilter() {
    if (this.c.edit) this.c.commitEdit()
    // Inside a table it is the table's Filter Button.
    if (this.c.tables.active()) return this.c.tables.setOption("filter", !this.c.tables.option("filter"))
    if (this.c.sheet.autoFilter) {
      this.setFilter({ filter: this.c.sheet.autoFilter, table: null }, null, "Filter")
      return
    }
    const range = this.filterRange()
    if (!range) {
      this.c.shell.alert("This can't be applied to the selected range. Select a single cell in a range and try again.")
      return
    }
    this.setFilter(null, { range, columns: {} }, "Filter")
  }

  /**
   * The filter's range grown down over rows typed under it since (Excel does this whenever a filter is applied). A
   * table's filter is the table's header and data rows.
   */
  private extended(host: FilterHost): Range {
    const range = host.filter.range
    if (host.table) return range
    const region = currentRegion(this.c.sheet, { row: range.r0, col: range.c0 })
    if (region.r0 > range.r0 || region.c0 > range.c0) return range
    return { ...range, r1: Math.max(range.r1, region.r1) }
  }

  /**
   * The ops that set a filter and show and hide its rows (null: the filter goes and every row it hid shows). `host`
   * is the filter being changed (null: a new sheet AutoFilter).
   */
  filterOps(host: FilterHost | null, filter: AutoFilter | null): Op[] {
    const sheet = this.c.sheet
    const old = host?.filter ?? null
    const shown = filter ?? (old ? { range: old.range, columns: {} } : null)
    const rows = shown ? filterRowEntries(sheet, this.c.wb.styles, shown, todaySerial()) : []
    const index = this.c.sheetIndex
    if (host?.table) {
      const table = host.table
      const tables = sheet.tables.map((t) => (t === table ? { ...copyTable(t), autoFilter: filter } : t))
      const ops: Op[] = [{ kind: "tables", sheet: index, tables }]
      if (rows.length) ops.push({ kind: "axis", sheet: index, axis: "rows", entries: rows })
      return ops
    }
    return [{ kind: "autoFilter", sheet: index, filter, ...(rows.length ? { rows } : {}) }]
  }

  private setFilter(host: FilterHost | null, filter: AutoFilter | null, label: string): boolean {
    const sel = this.c.selection
    return this.c.transact(this.filterOps(host, filter), { label, before: sel, after: sel }) !== null
  }

  /** A column's filter from its dropdown (null clears it). The range takes in rows added under it first. */
  filterColumn(col: number, column: ColumnFilter | null): boolean {
    const host = this.host()
    if (!host) return false
    if (this.c.edit) this.c.commitEdit()
    const f = host.filter
    const range = this.extended(host)
    const columns = { ...f.columns }
    if (column) columns[col - range.c0] = column
    else delete columns[col - range.c0]
    return this.setFilter(host, { ...f, range, columns }, "Filter")
  }

  /** Data › Clear: every column's filter goes (the dropdowns stay) and all rows show. */
  clearFilters(): boolean {
    const host = this.host()
    if (!host) return false
    const { sort: _sort, ...rest } = host.filter
    return this.setFilter(host, { ...rest, columns: {} }, "Clear")
  }

  /** Data › Reapply: the filters again, over the data as it is now. */
  reapply(): boolean {
    const host = this.host()
    if (!host) return false
    return this.setFilter(host, { ...host.filter, range: this.extended(host) }, "Reapply")
  }

  /** Whether any column of the filter has a filter (Clear and the status bar's count). */
  get filtering(): boolean {
    const f = this.autoFilter
    return Boolean(f && Object.keys(f.columns).length)
  }

  /** "3 of 10 records found": the filter's data rows shown, and all of them. Null when nothing is filtered. */
  records(): { shown: number; total: number } | null {
    const f = this.autoFilter
    if (!f || !this.filtering) return null
    const rows = this.c.sheet.rows
    let shown = 0
    for (let r = f.range.r0 + 1; r <= f.range.r1; r++) if (!rows.isHidden(r)) shown++
    return { shown, total: f.range.r1 - f.range.r0 }
  }

  /** Sort from a filter dropdown: the filter's data by that column; the button then shows the sort's arrow. */
  sortFiltered(col: number, descending: boolean, color?: { on: "cell" | "font"; color: string | null }): boolean {
    const host = this.host()
    if (!host) return false
    const f = host.filter
    if (this.c.edit) this.c.commitEdit()
    const range = this.extended(host)
    let ops: Op[]
    try {
      ops = sortOps(this.c.sheetIndex, this.c.sheet, this.c.wb.styles, range, {
        header: true,
        keys: [
          color
            ? { index: col, on: color.on === "cell" ? "cellColor" : "fontColor", color: color.color }
            : { index: col, descending },
        ],
      })
    } catch (error) {
      if (error instanceof SortError) {
        this.c.shell.alert(error.message)
        return false
      }
      throw error
    }
    const sorted = { ...f, range, sort: { col: col - range.c0, descending } }
    if (host.table) {
      const table = host.table
      const tables = this.c.sheet.tables.map((t) => (t === table ? { ...copyTable(t), autoFilter: sorted } : t))
      ops.push({ kind: "tables", sheet: this.c.sheetIndex, tables })
    } else ops.push({ kind: "autoFilter", sheet: this.c.sheetIndex, filter: sorted })
    const sel = this.c.selection
    return this.c.transact(ops, { label: "Sort", before: sel, after: sel }) !== null
  }
}
