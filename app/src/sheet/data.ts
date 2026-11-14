import {
  type AutoFilter,
  type ColumnFilter,
  currentRegion,
  filterRowEntries,
  type Op,
  guessHeader,
  isDateFormat,
  selectRange,
  SortError,
  sortOps,
  type Range,
  type SortOptions,
} from "@cells/core"
import { todaySerial, type SheetController } from "./controller"

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

  // ── AutoFilter ──

  /** Set by the grid: opens the dropdown of a filter column (Alt+Down on its header). */
  openMenu: ((col: number) => void) | null = null

  /** Alt+Down on a header of the filter: its dropdown. False when the active cell isn't one. */
  openMenuAtActive(): boolean {
    const f = this.c.sheet.autoFilter
    const { row, col } = this.c.selection.active
    if (!f || !this.openMenu || row !== f.range.r0 || col < f.range.c0 || col > f.range.c1) return false
    this.openMenu(col)
    return true
  }

  get autoFilter(): AutoFilter | null {
    return this.c.sheet.autoFilter
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
    if (this.c.sheet.autoFilter) {
      this.setFilter(null, "Filter")
      return
    }
    const range = this.filterRange()
    if (!range) {
      this.c.shell.alert("This can't be applied to the selected range. Select a single cell in a range and try again.")
      return
    }
    this.setFilter({ range, columns: {} }, "Filter")
  }

  /** The filter's range grown down over rows typed under it since (Excel does this whenever a filter is applied). */
  private extended(range: Range): Range {
    const region = currentRegion(this.c.sheet, { row: range.r0, col: range.c0 })
    if (region.r0 > range.r0 || region.c0 > range.c0) return range
    return { ...range, r1: Math.max(range.r1, region.r1) }
  }

  /** The ops that set `filter` and show and hide its rows (null: the filter goes and every row it hid shows). */
  private filterOps(filter: AutoFilter | null): Op[] {
    const sheet = this.c.sheet
    const shown = filter ?? (sheet.autoFilter ? { range: sheet.autoFilter.range, columns: {} } : null)
    const rows = shown ? filterRowEntries(sheet, this.c.wb.styles, shown, todaySerial()) : []
    return [{ kind: "autoFilter", sheet: this.c.sheetIndex, filter, ...(rows.length ? { rows } : {}) }]
  }

  private setFilter(filter: AutoFilter | null, label: string): boolean {
    const sel = this.c.selection
    return this.c.transact(this.filterOps(filter), { label, before: sel, after: sel }) !== null
  }

  /** A column's filter from its dropdown (null clears it). The range takes in rows added under it first. */
  filterColumn(col: number, column: ColumnFilter | null): boolean {
    const f = this.c.sheet.autoFilter
    if (!f) return false
    if (this.c.edit) this.c.commitEdit()
    const range = this.extended(f.range)
    const columns = { ...f.columns }
    if (column) columns[col - range.c0] = column
    else delete columns[col - range.c0]
    return this.setFilter({ ...f, range, columns }, "Filter")
  }

  /** Data › Clear: every column's filter goes (the dropdowns stay) and all rows show. */
  clearFilters(): boolean {
    const f = this.c.sheet.autoFilter
    if (!f) return false
    const { sort: _sort, ...rest } = f
    return this.setFilter({ ...rest, columns: {} }, "Clear")
  }

  /** Data › Reapply: the filters again, over the data as it is now. */
  reapply(): boolean {
    const f = this.c.sheet.autoFilter
    if (!f) return false
    return this.setFilter({ ...f, range: this.extended(f.range) }, "Reapply")
  }

  /** Whether any column of the filter has a filter (Clear and the status bar's count). */
  get filtering(): boolean {
    const f = this.c.sheet.autoFilter
    return Boolean(f && Object.keys(f.columns).length)
  }

  /** "3 of 10 records found": the filter's data rows shown, and all of them. Null when nothing is filtered. */
  records(): { shown: number; total: number } | null {
    const f = this.c.sheet.autoFilter
    if (!f || !this.filtering) return null
    const rows = this.c.sheet.rows
    let shown = 0
    for (let r = f.range.r0 + 1; r <= f.range.r1; r++) if (!rows.isHidden(r)) shown++
    return { shown, total: f.range.r1 - f.range.r0 }
  }

  /** Sort from a filter dropdown: the filter's data by that column; the button then shows the sort's arrow. */
  sortFiltered(col: number, descending: boolean, color?: { on: "cell" | "font"; color: string | null }): boolean {
    const f = this.c.sheet.autoFilter
    if (!f) return false
    if (this.c.edit) this.c.commitEdit()
    const range = this.extended(f.range)
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
    ops.push({
      kind: "autoFilter",
      sheet: this.c.sheetIndex,
      filter: { ...f, range, sort: { col: col - range.c0, descending } },
    })
    const sel = this.c.selection
    return this.c.transact(ops, { label: "Sort", before: sel, after: sel }) !== null
  }
}
