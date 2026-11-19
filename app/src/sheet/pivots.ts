import {
  chartRefRange,
  currentRegion,
  newPivotTable,
  pivotIsEmpty,
  pivotLayout,
  nextPivotCacheId,
  nextPivotName,
  pivotArea,
  pivotCacheFrom,
  quoteSheet,
  rangeName,
  Sheet,
  tableAt,
  innerFields,
  nextSheetName,
  nextTableId,
  nextTableName,
  pivotDetailRecords,
  outerField,
  displayText,
  pivotAt,
  pivotChoices,
  pivotFilterTarget,
  withFieldCleared,
  withFieldFilter,
  withPageItems,
  withPagePick,
  withShownItems,
  withSort,
  pivotCellAt,
  pivotCellsPlan,
  pivotOverwriteMessage,
  readPivotSource,
  refreshedCache,
  refreshedTable,
  withCollapsed,
  type CellWrite,
  type Op,
  type CellAddress,
  type PivotCache,
  type PivotCacheSource,
  type PivotCell,
  type PivotChoices,
  type PivotFilterSpec,
  type PivotFilterTarget,
  type PivotSort,
  type PivotTable,
  type Table,
} from "@cells/core"
import type { SheetController } from "./controller"

/** The width autofit adds for a cell's filter button (17 px and a gap). */
const FILTER_BUTTON_ROOM = 19

/** Excel's message when a new PivotTable would sit over another. */
export const PIVOT_OVERLAP_MESSAGE = "A PivotTable report cannot overlap another PivotTable report."

/** Where Create PivotTable puts the new PivotTable. */
export type PivotDestination = { kind: "newSheet" } | { kind: "existing"; sheet: number; at: CellAddress }

/** Which dialog a PivotTable filter dropdown opens, for the field at a filter button's cell. */
export type PivotFilterRequest = {
  kind: "label" | "value" | "top"
  row: number
  col: number
  field: number
  /** The operator picked in the submenu (Label and Value Filters). */
  op?: string
}

/**
 * PivotTables in the window: the one at the active cell, and Refresh (PivotTable Analyze › Refresh, Alt+F5, the cell
 * menu) and Refresh All (Data tab, Ctrl+Alt+F5), and expand / collapse (the +/− buttons, double-clicking an item,
 * the cell menu's Expand/Collapse), and the filter dropdowns (items ticked, sorts, report filters' picks), each one
 * undo step.
 */
export class PivotCommands {
  constructor(private readonly c: SheetController) {}

  /** The PivotTable at the active cell (the PivotTable tabs show while there is one). */
  active(): PivotTable | null {
    const { row, col } = this.c.selection.active
    return pivotAt(this.c.sheet, row, col)
  }

  any(): boolean {
    return this.c.wb.sheets.some((s) => s.pivots.length > 0)
  }

  /**
   * The ops that refresh these caches and lay out every PivotTable on them; a message when a source can't be read.
   * `ask` collects the sheets where a PivotTable would grow over other data.
   */
  refreshOps(cacheIds: readonly number[]): { ops: Op[]; overwrites: string[] } | string {
    const wb = this.c.wb
    const caches: PivotCache[] = [...wb.pivotCaches]
    const tables = new Map<number, PivotTable[]>(wb.sheets.map((s, i) => [i, [...s.pivots]]))
    const cellOps: Op[] = []
    const overwrites: string[] = []
    for (const id of cacheIds) {
      const index = caches.findIndex((c) => c.id === id)
      if (index < 0) continue
      const old = caches[index]
      const data = readPivotSource(wb.sheets, wb.styles, old.source)
      if (typeof data === "string") return data
      const next = refreshedCache(old, data)
      caches[index] = next
      wb.sheets.forEach((sheet, s) => {
        const list = tables.get(s)!
        list.forEach((t, k) => {
          if (t.cacheId !== id) return
          const plan = pivotCellsPlan(wb.sheets, wb.styles, s, t, refreshedTable(t, old, next), next)
          if (plan.overwrites && !overwrites.includes(sheet.name)) overwrites.push(sheet.name)
          cellOps.push(...plan.ops)
          list[k] = plan.table
        })
      })
    }
    const ops: Op[] = [{ kind: "pivotCaches", caches }, ...cellOps]
    tables.forEach((list, s) => {
      if (list.some((t, k) => t !== wb.sheets[s].pivots[k])) ops.push({ kind: "pivots", sheet: s, tables: list })
    })
    return { ops, overwrites }
  }

  private run(cacheIds: number[], label: string): boolean {
    const result = this.refreshOps(cacheIds)
    if (typeof result === "string") {
      this.c.shell.alert(result)
      return false
    }
    const go = () => {
      const entry = this.c.transact(result.ops, { label, before: this.c.selection, after: this.c.selection })
      if (entry) this.fitColumns(entry, result.ops)
    }
    if (result.overwrites.length)
      this.c.shell.confirm(pivotOverwriteMessage(this.c.printFile.fileName, result.overwrites[0]), go)
    else go()
    return true
  }

  /**
   * PivotTable Options › Refresh data when opening the file: those caches are read again as the workbook opens,
   * without an undo step (the file opens as it would in Excel). Sources that can't be read stay as saved.
   */
  refreshOnLoad(): number {
    const due = this.c.wb.pivotCaches.filter(
      (cache) => cache.refreshOnLoad && this.c.wb.sheets.some((s) => s.pivots.some((t) => t.cacheId === cache.id)),
    )
    if (!due.length) return 0
    const result = this.refreshOps(due.map((cache) => cache.id))
    if (typeof result === "string") return 0
    this.c.wb.transact(result.ops, { label: "Refresh", history: false })
    return due.length
  }

  /** Refresh: the active PivotTable's cache, and so every PivotTable that shares it. */
  refresh(): boolean {
    const t = this.active()
    return t ? this.run([t.cacheId], "Refresh") : false
  }

  /** Refresh All: every cache some PivotTable uses. */
  refreshAll(): boolean {
    const used = new Set(this.c.wb.sheets.flatMap((s) => s.pivots.map((t) => t.cacheId)))
    return used.size ? this.run([...used], "Refresh All") : false
  }

  /** The PivotTable at a cell, its cache, and the engine's cell there. */
  cellAt(row: number, col: number): { table: PivotTable; cache: PivotCache; cell: PivotCell | undefined } | null {
    const table = pivotAt(this.c.sheet, row, col)
    const cache = table && this.c.wb.pivotCaches.find((x) => x.id === table.cacheId)
    if (!table || !cache) return null
    return { table, cache, cell: pivotCellAt(table, cache, row, col) }
  }

  // ── Making and changing ──

  /**
   * Insert › PivotTable's starting source, as the dialog's Table/Range box shows it: the table at the active cell,
   * else the data around it (its current region), else nothing.
   */
  defaultSource(): string {
    const { row, col } = this.c.selection.active
    const table = tableAt(this.c.sheet, row, col)
    if (table) return table.name
    const sel = this.c.selection.ranges[0]
    const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
    const range = single ? currentRegion(this.c.sheet, { row, col }) : sel
    if (single && !this.c.sheet.cells.get(row, col) && range.r0 === range.r1 && range.c0 === range.c1) return ""
    return `${quoteSheet(this.c.sheet.name)}!${absolute(range)}`
  }

  /** What the Table/Range box says: a table or defined name, or a range on a sheet (this one when it names none). */
  parseSource(text: string): PivotCacheSource | null {
    const t = text.trim().replace(/^=/, "")
    if (!t) return null
    const wb = this.c.wb
    if (wb.sheets.some((s) => s.tables.some((x) => x.name.toLowerCase() === t.toLowerCase())))
      return {
        kind: "name",
        name: wb.sheets.flatMap((s) => s.tables).find((x) => x.name.toLowerCase() === t.toLowerCase())!.name,
      }
    if (wb.names.get(t)) return { kind: "name", name: wb.names.get(t)!.name }
    const ref = t.includes("!") ? t : `${quoteSheet(this.c.sheet.name)}!${t}`
    const at = chartRefRange(wb.sheets, ref)
    if (!at) return null
    return { kind: "range", sheet: at.sheet.name, ref: rangeName(at.range) }
  }

  /**
   * Create PivotTable's OK: a new cache from the source and an empty PivotTable, on a new sheet before this one (at
   * A3, as Excel puts it) or at a cell, shown and selected so the Fields pane opens on it. A message when it can't.
   */
  create(sourceText: string, dest: PivotDestination): string | null {
    const source = this.parseSource(sourceText)
    if (!source) return "Reference isn't valid."
    const wb = this.c.wb
    const data = readPivotSource(wb.sheets, wb.styles, source)
    if (typeof data === "string") return data
    if (!data.records.length && source.kind === "range") return "Reference isn't valid."
    const cache = pivotCacheFrom(nextPivotCacheId(wb.pivotCaches), source, data)
    const newSheet = dest.kind === "newSheet"
    const index = newSheet ? this.c.sheetIndex : dest.sheet
    const at = newSheet ? { row: 2, col: 0 } : dest.at
    const table = newPivotTable(nextPivotName(wb.sheets), cache, at)
    const sheets = [...wb.sheets]
    const name = nextSheetName(wb.sheets.map((s) => s.name))
    if (newSheet) sheets.splice(index, 0, new Sheet(name))
    else {
      const area = pivotArea({ ...table, location: { r0: at.row, c0: at.col, r1: at.row + 17, c1: at.col + 2 } })
      const hits = (r: { r0: number; c0: number; r1: number; c1: number }) =>
        r.r0 <= area.r1 && r.r1 >= area.r0 && r.c0 <= area.c1 && r.c1 >= area.c0
      if (sheets[index].pivots.some((p) => hits(pivotArea(p)))) return PIVOT_OVERLAP_MESSAGE
    }
    const plan = pivotCellsPlan(sheets, wb.styles, index, null, table, cache)
    const ops: Op[] = [
      ...(newSheet ? [{ kind: "addSheet" as const, sheet: index, name }] : []),
      { kind: "pivotCaches", caches: [...wb.pivotCaches, cache] },
      ...plan.ops,
      { kind: "pivots", sheet: index, tables: [...sheets[index].pivots, plan.table] },
    ]
    const go = () => {
      if (this.c.edit) this.c.commitEdit()
      const entry = this.c.transact(ops, { label: "Insert PivotTable", before: this.c.selection, sheetAfter: index })
      if (!entry) return
      if (index !== this.c.sheetIndex || newSheet) this.c.activate(index)
      this.fitColumns(entry, ops)
      this.c.goTo({ r0: at.row, c0: at.col, r1: at.row, c1: at.col })
      this.fieldsOpen = true
      this.fieldsAsked = true
      this.c.changed()
    }
    if (plan.overwrites) this.c.shell.confirm(pivotOverwriteMessage(this.c.printFile.fileName, sheets[index].name), go)
    else go()
    return null
  }

  /**
   * The PivotTable Fields pane is on (Analyze › Field List): it shows while the active cell is in a PivotTable. On a
   * phone it covers the screen, so there it shows only once asked for (a new PivotTable, Field List).
   */
  fieldsOpen = true
  fieldsAsked = false

  setFieldsOpen(open: boolean) {
    this.fieldsOpen = open
    this.fieldsAsked = open
    this.c.changed()
  }

  /** Whether the Fields pane shows now, in the window's layout. */
  fieldsShown(phone: boolean): boolean {
    return this.fieldsOpen && (!phone || this.fieldsAsked) && this.active() !== null
  }

  /** The active PivotTable and its cache, for the Fields pane. */
  activeWithCache(): { table: PivotTable; cache: PivotCache } | null {
    const table = this.active()
    const cache = table && this.c.wb.pivotCaches.find((x) => x.id === table.cacheId)
    return table && cache ? { table, cache } : null
  }

  /** A change to the active PivotTable (the Fields pane, the field menus), laid out again as one undo step. */
  change(make: (table: PivotTable, cache: PivotCache) => PivotTable, label = "PivotTable Field"): boolean {
    const at = this.activeWithCache()
    return at ? this.relayout(at, make(at.table, at.cache), label) : false
  }

  /**
   * Autofit column widths on update: the columns of each PivotTable the ops lay out on the active sheet fit its cells
   * (wider or narrower, as in Excel), in the same undo step. An empty PivotTable keeps the widths it has.
   */
  private fitColumns(entry: Parameters<SheetController["wb"]["amend"]>[0], ops: readonly Op[]) {
    const fit = this.c.columnFit
    const s = this.c.sheetIndex
    if (!fit) return
    const sizes = new Map<number, number>()
    for (const op of ops) {
      if (op.kind !== "pivots" || op.sheet !== s) continue
      for (const t of op.tables) {
        if (t.autoFit === false || pivotIsEmpty(t)) continue
        const area = pivotArea(t)
        const grow = (c: number, size: number | null) => {
          if (size !== null) sizes.set(c, Math.max(sizes.get(c) ?? 0, size))
        }
        for (let c = area.c0; c <= area.c1; c++) grow(c, fit(c, { r0: area.r0, r1: area.r1 }))
        // A cell with a filter button gets room for it, as Excel's autofit gives.
        const cache = this.c.wb.pivotCaches.find((x) => x.id === t.cacheId)
        if (cache)
          for (const cell of pivotLayout(t, cache).cells) {
            if (!pivotFilterTarget(t, cache, cell)) continue
            const size = fit(cell.col, { r0: cell.row, r1: cell.row })
            if (size !== null) grow(cell.col, size + FILTER_BUTTON_ROOM)
          }
      }
    }
    const cols = this.c.sheet.cols
    const entries = [...sizes]
      .filter(([c, size]) => Math.abs(size - cols.size(c)) >= 1)
      .map(([index, size]) => ({ index, size, hidden: false }))
    if (entries.length) this.c.wb.amend(entry, [{ kind: "axis", sheet: s, axis: "cols", entries }])
  }

  // ── Show Details ──

  /** The cell is a value of a PivotTable whose details can be shown (the cell menu's Show Details). */
  canShowDetails(row: number, col: number): boolean {
    return Boolean(this.cellAt(row, col)?.cell?.lines)
  }

  /**
   * Show Details (double-clicking a value): the source rows it sums up, as a table on a new sheet before the
   * PivotTable's (Excel's SheetN, TableN in TableStyleMedium2, dates in their format, columns fitted), as one undo
   * step. Rows come in the source's order (Excel's own order follows its internal index). False off a value.
   */
  showDetails(row: number, col: number): boolean {
    const at = this.cellAt(row, col)
    if (!at?.cell?.lines) return false
    if (at.table.enableDrill === false) {
      this.c.shell.alert("Cannot change this part of a PivotTable report.")
      return true
    }
    const records = pivotDetailRecords(at.table, at.cache, at.cell)
    if (!records) return false
    const wb = this.c.wb
    const index = this.c.sheetIndex
    const name = nextSheetName(wb.sheets.map((s) => s.name))
    const fields = at.cache.fields
    const styleOf = fields.map((f) =>
      f.numFmt && f.numFmt !== "General" ? wb.styles.intern({ ...wb.styles.get(0), numberFormat: f.numFmt }) : 0,
    )
    const cells: CellWrite[] = fields.map((f, col) => ({ row: 0, col, data: { value: f.name } }))
    records.forEach((r, i) =>
      at.cache.records[r].forEach((value, col) => {
        if (value === null || value === "") return
        cells.push({ row: i + 1, col, data: { value, ...(styleOf[col] ? { style: styleOf[col] } : {}) } })
      }),
    )
    const range = { r0: 0, c0: 0, r1: records.length, c1: fields.length - 1 }
    const table: Table = {
      id: nextTableId(wb.sheets),
      name: nextTableName(
        wb.sheets,
        wb.names.all().map((n) => n.name),
      ),
      range,
      header: true,
      totals: false,
      columns: fields.map((f) => ({ name: f.name })),
      style: "TableStyleMedium2",
      firstColumn: false,
      lastColumn: false,
      rowStripes: true,
      columnStripes: false,
      autoFilter: { range: { ...range }, columns: {} },
    }
    const ops: Op[] = [
      { kind: "addSheet", sheet: index, name },
      { kind: "setCells", sheet: index, cells },
      { kind: "tables", sheet: index, tables: [table] },
    ]
    if (this.c.edit) this.c.commitEdit()
    const entry = this.c.transact(ops, { label: "Show Details", before: this.c.selection, sheetAfter: index })
    if (!entry) return true
    this.c.activate(index)
    const fit = this.c.columnFit
    const widths = []
    for (let c = 0; fit && c < fields.length; c++) {
      const size = fit(c, { r0: range.r0, r1: range.r1 })
      if (size !== null && size > this.c.sheet.cols.size(c)) widths.push({ index: c, size, hidden: false })
    }
    if (widths.length) wb.amend(entry, [{ kind: "axis", sheet: index, axis: "cols", entries: widths }])
    this.c.goTo(range)
    return true
  }

  /** What the cell menu's Expand/Collapse can do at the active cell (null: not on a row or column label). */
  drillState(): { expand: boolean; collapse: boolean; expandField: boolean; collapseField: boolean } | null {
    const { row, col } = this.c.selection.active
    const at = this.cellAt(row, col)
    const item = at?.cell?.item
    if (!at || !item) return null
    const f = at.table.fields[item.field]
    const own = f.items.find((i) => i.x === item.x)
    const inner = innerFields(at.table, item).length > 0
    const fieldOf = inner ? item.field : outerField(at.table, item)
    const anyCollapsed = (field: number) => at.table.fields[field].items.some((i) => i.collapsed)
    const anyOpen = (field: number) =>
      at.table.fields[field].items.length < (at.cache.fields[field].items?.length ?? 0) ||
      at.table.fields[field].items.some((i) => !i.collapsed)
    return {
      expand: inner && Boolean(own?.collapsed),
      collapse: inner && !own?.collapsed,
      expandField: inner && anyCollapsed(item.field),
      collapseField: fieldOf !== null && anyOpen(fieldOf),
    }
  }

  /** The +/− button of the label at a cell, or double-clicking it: the item opens or closes. */
  toggle(row: number, col: number): boolean {
    const at = this.cellAt(row, col)
    const cell = at?.cell
    if (!at || !cell?.item || !cell.button) return false
    const collapse = cell.button === "collapse"
    return this.relayout(
      at,
      withCollapsed(at.table, at.cache, cell.item.field, [cell.item.x], collapse),
      collapse ? "Collapse" : "Expand",
    )
  }

  /**
   * The cell menu's Expand/Collapse at the active cell: the item, or every item of its field. Collapse Entire Field on
   * the innermost field closes the field above it, as in Excel.
   */
  drill(scope: "item" | "field", collapse: boolean): boolean {
    const { row, col } = this.c.selection.active
    const at = this.cellAt(row, col)
    const item = at?.cell?.item
    if (!at || !item) return false
    const inner = innerFields(at.table, item).length > 0
    const label = collapse ? "Collapse" : "Expand"
    if (scope === "item")
      return inner && this.relayout(at, withCollapsed(at.table, at.cache, item.field, [item.x], collapse), label)
    const field = inner ? item.field : collapse ? outerField(at.table, item) : null
    return field !== null && this.relayout(at, withCollapsed(at.table, at.cache, field, null, collapse), label)
  }

  // ── Filter dropdowns ──

  /** Set by the grid while it shows the buttons: opens the dropdown of the button at a cell. */
  openMenu: ((row: number, col: number) => void) | null = null

  /** Alt+Down on a cell with a PivotTable filter button: its dropdown. */
  openMenuAtActive(): boolean {
    const { row, col } = this.c.selection.active
    if (!this.openMenu || !this.filterTargetAt(row, col)) return false
    this.openMenu(row, col)
    return true
  }

  /** The filter button a cell shows (Row Labels / Column Labels, a field's header, a report filter's value). */
  filterTargetAt(row: number, col: number): PivotFilterTarget | null {
    const at = this.cellAt(row, col)
    return at ? pivotFilterTarget(at.table, at.cache, at.cell) : null
  }

  /** A dropdown's list for one of its fields, labels in the field's number format. */
  choices(row: number, col: number, field: number): PivotChoices | null {
    const at = this.cellAt(row, col)
    if (!at || !pivotFilterTarget(at.table, at.cache, at.cell)?.fields.includes(field)) return null
    return pivotChoices(at.table, at.cache, field, (v, numFmt) => displayText(v, numFmt ?? "General"))
  }

  /** The dropdown's OK: only `shown` ticked (null: Clear Filter), one undo step. */
  showItems(row: number, col: number, field: number, shown: ReadonlySet<number> | null): boolean {
    const at = this.cellAt(row, col)
    if (!at) return false
    const page = at.table.pages.some((p) => p.field === field)
    const next =
      page && shown ? withPageItems(at.table, at.cache, field, shown) : withShownItems(at.table, at.cache, field, shown)
    return this.relayout(at, next, shown ? "Filter" : "Clear Filter")
  }

  /** A Label Filter, Value Filter or Top 10 on a field (null: removed), from the submenus and their dialogs. */
  setFilter(row: number, col: number, field: number, spec: PivotFilterSpec | null): boolean {
    const at = this.cellAt(row, col)
    return at ? this.relayout(at, withFieldFilter(at.table, field, spec), spec ? "Filter" : "Clear Filter") : false
  }

  /** Clear Filter From "Field": ticked items and the field's filter both. */
  clearField(row: number, col: number, field: number): boolean {
    const at = this.cellAt(row, col)
    return at ? this.relayout(at, withFieldCleared(at.table, at.cache, field), "Clear Filter") : false
  }

  /** The value fields a Value Filter or Top 10 can judge by. */
  valueNames(row: number, col: number): string[] {
    return this.cellAt(row, col)?.table.values.map((v) => v.name) ?? []
  }

  /** A report filter set to one item, or to (All) (null). */
  pickPage(row: number, col: number, field: number, x: number | null): boolean {
    const at = this.cellAt(row, col)
    return at ? this.relayout(at, withPagePick(at.table, at.cache, field, x), "Filter") : false
  }

  /** Sort A to Z / Z to A on a field. */
  sortField(row: number, col: number, field: number, sort: PivotSort): boolean {
    const at = this.cellAt(row, col)
    return at ? this.relayout(at, withSort(at.table, at.cache, field, sort), "Sort") : false
  }

  private relayout(at: { table: PivotTable; cache: PivotCache }, next: PivotTable, label: string): boolean {
    if (next === at.table) return false
    const s = this.c.wb.sheets.indexOf(this.c.sheet)
    const plan = pivotCellsPlan(this.c.wb.sheets, this.c.wb.styles, s, at.table, next, at.cache)
    const tables = this.c.sheet.pivots.map((t) => (t === at.table ? plan.table : t))
    const ops: Op[] = [...plan.ops, { kind: "pivots", sheet: s, tables }]
    const go = () => {
      const entry = this.c.transact(ops, { label, before: this.c.selection })
      if (entry) this.fitColumns(entry, ops)
      // The active cell stays in the PivotTable (at its top left) when the change shrinks the table away from it, so
      // the Fields pane stays open on it.
      const { row, col } = this.c.selection.active
      const area = pivotArea(plan.table)
      if (row < area.r0 || row > area.r1 || col < area.c0 || col > area.c1)
        this.c.goTo({ r0: plan.table.at.row, c0: plan.table.at.col, r1: plan.table.at.row, c1: plan.table.at.col })
    }
    if (plan.overwrites) this.c.shell.confirm(pivotOverwriteMessage(this.c.printFile.fileName, this.c.sheet.name), go)
    else go()
    return true
  }
}

/** A range with $ signs, as Excel's Table/Range box shows it. */
function absolute(r: { r0: number; c0: number; r1: number; c1: number }): string {
  return rangeName(r)
    .split(":")
    .map((part) => part.replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2"))
    .join(":")
}
