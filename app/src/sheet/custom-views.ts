import {
  copyPageSetup,
  hiddenSpans,
  newViewGuid,
  type AxisEntry,
  type CustomSheetView,
  type CustomView,
  type Op,
  type Sheet,
} from "@cells/core"
import type { SheetView as FileSheetView } from "@cells/io"
import type { SheetController } from "./controller"

/** Excel's message before a view is replaced or deleted (Cells' wording; see verify-list). */
export const VIEW_MESSAGES = {
  replace: (name: string) => `A view named '${name}' already exists. Do you want to replace it?`,
  remove: (name: string) => `Delete the view '${name}'?`,
}

/**
 * View › Custom Views: Add saves how every sheet looks now (and, if chosen, its hidden rows and columns and print
 * settings) under a name; Show puts it all back. Views aren't undone (as in Excel); showing one that changes hidden
 * rows or print settings is one undo step for those.
 */
export class CustomViewCommands {
  constructor(private readonly c: SheetController) {}

  list(): readonly CustomView[] {
    return this.c.wb.customViews
  }

  find(name: string): CustomView | undefined {
    const key = name.trim().toLowerCase()
    return this.c.wb.customViews.find((v) => v.name.toLowerCase() === key)
  }

  /** Adds (or replaces) a view of the workbook as it looks now. */
  add(name: string, printSettings: boolean, hiddenRowsCols: boolean) {
    const c = this.c
    if (c.edit) c.commitEdit()
    const files = c.saveViews()
    const sheets = new Map<Sheet, CustomSheetView>()
    c.wb.sheets.forEach((sheet, i) => {
      const v = files[i]
      sheets.set(sheet, {
        zoom: v.zoom,
        topLeft: { ...v.topLeft },
        active: { ...v.active },
        selection: v.selection.map((r) => ({ ...r })),
        freeze: v.freeze ? { ...v.freeze } : null,
        showGridlines: v.showGridlines,
        showHeadings: v.showHeadings,
        mode: v.mode ?? "normal",
        hiddenRows: sheet.rows.anyHidden ? hiddenSpans((r) => sheet.rows.isHidden(r), sheet.rows.count, Infinity) : [],
        hiddenCols: sheet.cols.anyHidden ? hiddenSpans((k) => sheet.cols.isHidden(k), sheet.cols.count, Infinity) : [],
        pageSetup: copyPageSetup(sheet.pageSetup),
      })
    })
    const old = this.find(name)
    const view: CustomView = {
      name: name.trim(),
      guid: old?.guid ?? newViewGuid(),
      printSettings,
      hiddenRowsCols,
      activeSheet: c.sheet,
      sheets,
    }
    c.wb.customViews = old ? c.wb.customViews.map((v) => (v === old ? view : v)) : [...c.wb.customViews, view]
    c.unsavedChange()
  }

  remove(name: string) {
    const view = this.find(name)
    if (!view) return
    this.c.wb.customViews = this.c.wb.customViews.filter((v) => v !== view)
    this.c.unsavedChange()
  }

  /** Shows a view: hidden rows and columns and print settings first (one undo step), then each sheet's look. */
  show(name: string) {
    const c = this.c
    const view = this.find(name)
    if (!view) return
    if (c.edit) c.commitEdit()
    c.group.clear(false)
    const ops: Op[] = []
    c.wb.sheets.forEach((sheet, index) => {
      const v = view.sheets.get(sheet)
      if (!v) return
      if (view.hiddenRowsCols) {
        for (const axis of ["rows", "cols"] as const) {
          const entries = hiddenChanges(sheet, axis, axis === "rows" ? v.hiddenRows : v.hiddenCols)
          if (entries.length) ops.push({ kind: "axis", sheet: index, axis, entries })
        }
      }
      if (view.printSettings && v.pageSetup && JSON.stringify(v.pageSetup) !== JSON.stringify(sheet.pageSetup))
        ops.push({ kind: "pageSetup", sheet: index, setup: copyPageSetup(v.pageSetup) })
    })
    if (ops.length) c.transact(ops, { label: "Show View", before: c.selection, after: c.selection, perSheet: true })
    const now = c.saveViews()
    const files: FileSheetView[] = c.wb.sheets.map((sheet, i) => {
      const v = view.sheets.get(sheet)
      const base = { ...now[i], selected: false }
      if (!v) return base
      return {
        ...base,
        zoom: v.zoom,
        mode: v.mode,
        zooms: { ...base.zooms, [v.mode]: v.zoom },
        topLeft: { ...v.topLeft },
        active: { ...v.active },
        selection: v.selection.map((r) => ({ ...r })),
        freeze: v.freeze ? { ...v.freeze } : null,
        split: null,
        showGridlines: v.showGridlines,
        showHeadings: v.showHeadings,
      }
    })
    const active = view.activeSheet ? c.wb.sheets.indexOf(view.activeSheet) : c.sheetIndex
    c.loadViews(files, active >= 0 ? active : c.sheetIndex)
  }
}

/** The axis entries that make a sheet's hidden lines match `spans` (only the lines that change). */
function hiddenChanges(sheet: Sheet, axis: "rows" | "cols", spans: readonly [number, number][]): AxisEntry[] {
  const a = axis === "rows" ? sheet.rows : sheet.cols
  const want = new Set<number>()
  for (const [from, to] of spans) for (let i = from; i <= to && i < a.count; i++) want.add(i)
  const lines = new Set(want)
  if (a.anyHidden)
    for (const [from, to] of hiddenSpans((i) => a.isHidden(i), a.count, Infinity))
      for (let i = from; i <= to; i++) lines.add(i)
  const out: AxisEntry[] = []
  for (const index of [...lines].sort((x, y) => x - y)) {
    const hidden = want.has(index)
    if (hidden === a.isHidden(index)) continue
    out.push({ index, size: a.hasCustomSize(index) ? a.storedSize(index) : null, hidden })
  }
  return out
}
