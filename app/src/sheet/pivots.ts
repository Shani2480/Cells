import {
  innerFields,
  outerField,
  pivotAt,
  pivotCellAt,
  pivotCellsPlan,
  pivotOverwriteMessage,
  readPivotSource,
  refreshedCache,
  refreshedTable,
  withCollapsed,
  type Op,
  type PivotCache,
  type PivotCell,
  type PivotTable,
} from "@cells/core"
import type { SheetController } from "./controller"

/**
 * PivotTables in the window: the one at the active cell, and Refresh (PivotTable Analyze › Refresh, Alt+F5, the cell
 * menu) and Refresh All (Data tab, Ctrl+Alt+F5), and expand / collapse (the +/− buttons, double-clicking an item,
 * the cell menu's Expand/Collapse), each one undo step.
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
      this.c.transact(result.ops, { label, before: this.c.selection, after: this.c.selection })
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
    return this.relayout(at, withCollapsed(at.table, at.cache, cell.item.field, [cell.item.x], collapse), collapse)
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
    if (scope === "item")
      return inner && this.relayout(at, withCollapsed(at.table, at.cache, item.field, [item.x], collapse), collapse)
    const field = inner ? item.field : collapse ? outerField(at.table, item) : null
    return field !== null && this.relayout(at, withCollapsed(at.table, at.cache, field, null, collapse), collapse)
  }

  private relayout(at: { table: PivotTable; cache: PivotCache }, next: PivotTable, collapse: boolean): boolean {
    if (next === at.table) return false
    const s = this.c.wb.sheets.indexOf(this.c.sheet)
    const plan = pivotCellsPlan(this.c.wb.sheets, this.c.wb.styles, s, at.table, next, at.cache)
    const tables = this.c.sheet.pivots.map((t) => (t === at.table ? plan.table : t))
    const ops: Op[] = [...plan.ops, { kind: "pivots", sheet: s, tables }]
    const go = () => {
      this.c.transact(ops, { label: collapse ? "Collapse" : "Expand", before: this.c.selection })
    }
    if (plan.overwrites) this.c.shell.confirm(pivotOverwriteMessage(this.c.printFile.fileName, this.c.sheet.name), go)
    else go()
    return true
  }
}
