import {
  pivotAt,
  pivotCellsPlan,
  pivotOverwriteMessage,
  readPivotSource,
  refreshedCache,
  refreshedTable,
  type Op,
  type PivotCache,
  type PivotTable,
} from "@cells/core"
import type { SheetController } from "./controller"

/**
 * PivotTables in the window: the one at the active cell, and Refresh (PivotTable Analyze › Refresh, Alt+F5, the cell
 * menu) and Refresh All (Data tab, Ctrl+Alt+F5), each one undo step.
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
}
