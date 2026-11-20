import {
  DEFAULT_SLICER_STYLE,
  isChart,
  nextObjectId,
  SLICER_BUTTON_HEIGHT,
  SLICER_SIZE,
  SLICER_STYLES,
  slicerFilter,
  slicerItems,
  slicerTarget,
  pivotSlicerFiltered,
  pivotSlicerItems,
  pivotSlicerShown,
  pivotSlicerTarget,
  tableFilterRange,
  type AutoFilter,
  type Op,
  type PivotCache,
  type SheetObject,
  type SheetSlicer,
} from "@cells/core"
import { setSlicerSource, type SlicerView } from "@/drawing/slicer"
import { anchorOf } from "./charts"
import { todaySerial, type SheetController } from "./controller"

export type { SlicerView }

/** Slicer Settings: what the dialog changes. */
export type SlicerSettings = Pick<
  SheetSlicer,
  "name" | "caption" | "header" | "sort" | "customLists" | "hideNoData" | "crossFilter"
>

/** Room around the buttons: 9 px each side, 3 px between columns (Excel's layout). */
const SIDES = 18
const GAP = 3

/** How a click picks: this button only, this one on or off (Ctrl, Multi-Select), or a run of them (Shift). */
export type SlicerPick = "only" | "toggle" | "range"

/**
 * Slicers on the sheet: their buttons (from the table they filter) and picking them. A pick is the table's filter on
 * the slicer's column, set as Excel does, so it undoes as one step and the Filter Button's dropdown agrees.
 */
export class SlicerCommands {
  /** Slicers with Multi-Select on (the header's button), by id: a view state, not saved. */
  private readonly multi = new Set<number>()
  /** The button a Shift+click runs from, per slicer. */
  private readonly anchor = new Map<number, string>()
  private cache: { version: number; views: Map<number, SlicerView> } | null = null

  constructor(private readonly c: SheetController) {
    setSlicerSource((o) => this.view(o))
  }

  /** The selected slicer (the Slicer tab's), or null. */
  active(): SheetSlicer | null {
    const list = this.c.objects.drawings()
    return list.length === 1 && list[0].kind === "slicer" ? list[0] : null
  }

  /** The selected slicers (the Slicer tab changes them all). */
  selected(): SheetSlicer[] {
    return this.c.objects.drawings().filter((o): o is SheetSlicer => o.kind === "slicer")
  }

  view(o: SheetSlicer): SlicerView {
    if (!this.cache || this.cache.version !== this.c.wb.version)
      this.cache = { version: this.c.wb.version, views: new Map() }
    let v = this.cache.views.get(o.id)
    if (!v && o.pivot) {
      const target = pivotSlicerTarget(this.c.wb.sheets, this.c.wb.pivotCaches, o)
      v = {
        items: target ? pivotSlicerItems(target, o) : [],
        filtered: target ? pivotSlicerFiltered(target) : false,
        multi: false,
      }
      this.cache.views.set(o.id, v)
    }
    if (!v) {
      const target = slicerTarget(this.c.wb.sheets, o)
      const look = target?.sheet === this.c.sheetIndex ? this.c.data.look : undefined
      const filter = target?.table.autoFilter
      const offset = target && filter ? target.col - filter.range.c0 : -1
      v = {
        items: slicerItems(this.c.wb.sheets, this.c.wb.styles, o, todaySerial(), look),
        filtered: Boolean(filter && filter.columns[offset]),
        multi: false,
      }
      this.cache.views.set(o.id, v)
    }
    return { ...v, multi: this.multi.has(o.id) }
  }

  isMulti(o: SheetSlicer): boolean {
    return this.multi.has(o.id)
  }

  /** The header's Multi-Select button (Alt+S): clicks then pick buttons on and off. */
  toggleMulti(o: SheetSlicer) {
    if (this.multi.has(o.id)) this.multi.delete(o.id)
    else this.multi.add(o.id)
    this.c.changed()
  }

  /** A click on a button (`key`, as slicerItems gives it). */
  pick(o: SheetSlicer, key: string, how: SlicerPick): boolean {
    const items = this.view(o).items
    let picked = new Set(items.filter((i) => i.selected).map((i) => i.key))
    const from = this.anchor.get(o.id)
    if (how === "range" && from !== undefined && items.some((i) => i.key === from)) {
      const a = items.findIndex((i) => i.key === from)
      const b = items.findIndex((i) => i.key === key)
      picked = new Set(items.slice(Math.min(a, b), Math.max(a, b) + 1).map((i) => i.key))
    } else if (how === "toggle" || this.multi.has(o.id)) {
      // With nothing filtered, the first toggle leaves just that one out, as in Excel.
      if (picked.has(key)) {
        picked.delete(key)
        // Excel keeps at least one button picked.
        if (!picked.size) return false
      } else picked.add(key)
      this.anchor.set(o.id, key)
    } else {
      picked = new Set([key])
      this.anchor.set(o.id, key)
    }
    if (o.pivot) return this.applyPivot(o, pivotSlicerShown(items, picked))
    return this.apply(o, slicerFilter(items, picked))
  }

  /** The header's Clear Filter button (Alt+C). */
  clear(o: SheetSlicer): boolean {
    this.anchor.delete(o.id)
    if (o.pivot) return this.applyPivot(o, null)
    return this.apply(o, null)
  }

  /** A PivotTable slicer's pick: the field's items shown in each PivotTable it is connected to. */
  private applyPivot(o: SheetSlicer, shown: Set<number> | null): boolean {
    const target = pivotSlicerTarget(this.c.wb.sheets, this.c.wb.pivotCaches, o)
    if (!target) return false
    const changes = target.tables.map((t) => ({ ...t, field: target.field, shown }))
    return this.c.pivots.slicerPick(changes, target.cache, "Slicer")
  }

  /** The items a PivotTable slicer shows now (null: all), to give a PivotTable newly connected to it. */
  private pivotShown(o: SheetSlicer): Set<number> | null {
    const items = this.view(o).items
    return pivotSlicerShown(items, new Set(items.filter((i) => i.selected).map((i) => i.key)))
  }

  // ── Report Connections (Slicer tab) and Filter Connections (PivotTable Analyze) ──

  /** The PivotTables a PivotTable slicer can filter (those of its cache), ticked when connected. */
  reportConnections(o: SheetSlicer): { key: string; name: string; sheet: string; on: boolean }[] {
    const link = o.pivot
    if (!link) return []
    const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase()
    return this.c.wb.sheets.flatMap((s) =>
      s.pivots
        .filter((t) => t.cacheId === link.cacheId)
        .map((t) => ({
          key: `${s.name}!${t.name}`,
          name: t.name,
          sheet: s.name,
          on: link.tables.some((x) => same(x.sheet, s.name) && x.name === t.name),
        })),
    )
  }

  /** Report Connections' OK: the slicer filters the PivotTables `keys` ("Sheet!Name"); new ones take its picks. */
  setReportConnections(o: SheetSlicer, keys: readonly string[]): string | null {
    if (!o.pivot) return null
    const list = this.reportConnections(o)
    const picked = list.filter((x) => keys.includes(x.key))
    if (!picked.length) return "Select at least one PivotTable for the slicer to filter."
    const tables = picked.map((x) => ({ sheet: x.sheet, name: x.name }))
    this.connect([{ slicer: o, tables }], "Report Connections")
    return null
  }

  /** The PivotTable slicers the active PivotTable can be filtered by (its cache's), ticked when connected. */
  filterConnections(): { id: number; name: string; caption: string; sheet: string; on: boolean }[] {
    const at = this.c.pivots.active()
    if (!at) return []
    const here = this.c.sheet.name.toLowerCase()
    return this.all().flatMap(({ sheet, slicer }) =>
      slicer.pivot && slicer.pivot.cacheId === at.cacheId
        ? [
            {
              id: slicer.id,
              name: slicer.name,
              caption: slicer.caption,
              sheet: this.c.wb.sheets[sheet].name,
              on: slicer.pivot.tables.some((t) => t.sheet.toLowerCase() === here && t.name === at.name),
            },
          ]
        : [],
    )
  }

  /** Filter Connections' OK: the active PivotTable is filtered by the slicers `ids` (and no other of its cache's). */
  setFilterConnections(ids: readonly number[]): boolean {
    const at = this.c.pivots.active()
    if (!at) return false
    const here = { sheet: this.c.sheet.name, name: at.name }
    const same = (t: { sheet: string; name: string }) =>
      t.sheet.toLowerCase() === here.sheet.toLowerCase() && t.name === here.name
    const changes = this.filterConnections().flatMap((x) => {
      const slicer = this.all().find((s) => s.slicer.id === x.id)!.slicer
      const want = ids.includes(x.id)
      if (want === x.on) return []
      const others = slicer.pivot!.tables.filter((t) => !same(t))
      // A slicer keeps at least one PivotTable (as Excel's Report Connections asks).
      if (!want && !others.length) return []
      return [{ slicer, tables: want ? [...others, here] : others }]
    })
    return this.connect(changes, "Filter Connections")
  }

  /**
   * Slicers connected to these PivotTables, as one undo step: each PivotTable newly connected takes the slicer's
   * picks, and one let go shows all the field's items again, as in Excel.
   */
  private connect(
    changes: readonly { slicer: SheetSlicer; tables: { sheet: string; name: string }[] }[],
    label: string,
  ) {
    if (!changes.length) return false
    const sheets = this.c.wb.sheets
    const extra: Op[] = []
    const picks: Parameters<SheetController["pivots"]["slicerPick"]>[0][number][] = []
    let cache: PivotCache | null = null
    const bySheet = new Map<number, SheetObject[]>()
    for (const { slicer, tables } of changes) {
      const target = pivotSlicerTarget(sheets, this.c.wb.pivotCaches, slicer)
      const shown = this.pivotShown(slicer)
      const had = (t: { sheet: string; name: string }) =>
        slicer.pivot!.tables.some((x) => x.sheet.toLowerCase() === t.sheet.toLowerCase() && x.name === t.name)
      const where = this.all().find((x) => x.slicer.id === slicer.id)!.sheet
      const list = bySheet.get(where) ?? [...sheets[where].objects]
      bySheet.set(
        where,
        list.map((x) => (x.id === slicer.id && !isChart(x) ? { ...slicer, pivot: { ...slicer.pivot!, tables } } : x)),
      )
      if (!target) continue
      cache = target.cache
      const kept = (t: { sheet: string; name: string }) =>
        tables.some((x) => x.sheet.toLowerCase() === t.sheet.toLowerCase() && x.name === t.name)
      const pick = (t: { sheet: string; name: string }, items: Set<number> | null) => {
        const sheet = sheets.findIndex((s) => s.name.toLowerCase() === t.sheet.toLowerCase())
        const table = sheets[sheet]?.pivots.find((p) => p.name === t.name)
        if (table) picks.push({ sheet, table, field: target.field, shown: items })
      }
      for (const t of tables) if (!had(t)) pick(t, shown)
      for (const t of slicer.pivot!.tables) if (!kept(t)) pick(t, null)
    }
    for (const [sheet, objects] of bySheet) extra.push({ kind: "objects", sheet, objects })
    if (!cache) return this.c.transact(extra, { label, before: this.c.selection, after: this.c.selection }) !== null
    return this.c.pivots.slicerPick(picks, cache, label, extra)
  }

  private apply(o: SheetSlicer, column: ReturnType<typeof slicerFilter>): boolean {
    const target = slicerTarget(this.c.wb.sheets, o)
    if (!target) return false
    if (this.c.edit) this.c.commitEdit()
    const { table } = target
    const before: AutoFilter = table.autoFilter ?? { range: tableFilterRange(table), columns: {} }
    const offset = target.col - before.range.c0
    const columns = { ...before.columns }
    if (column) columns[offset] = column
    else delete columns[offset]
    if (!column && !before.columns[offset]) return false
    const filter: AutoFilter = { ...before, columns }
    const ops = this.c.data.filterOps({ filter: before, table }, filter, target.sheet)
    const sel = this.c.selection
    return this.c.transact(ops, { label: "Slicer", before: sel, after: sel }) !== null
  }

  /** Scrolls a slicer's buttons (the mouse wheel over it): a view change, kept when saved, not undone. */
  scroll(o: SheetSlicer, rows: number, shownRows: number, totalRows: number) {
    const columns = Math.max(1, o.columns)
    const first = Math.max(0, Math.min(totalRows - shownRows, Math.floor((o.startItem ?? 0) / columns) + rows))
    const next = first * columns
    if (next === (o.startItem ?? 0)) return
    if (next) o.startItem = next
    else delete o.startItem
    this.c.changed()
  }

  // ── Insert › Slicer and the Slicer tab ──

  /**
   * Insert › Slicer, Table Design › Insert Slicer, PivotTable Analyze › Insert Slicer: the Insert Slicers dialog for
   * the active cell's table or PivotTable.
   */
  openInsert() {
    if (!this.c.tables.active() && !this.c.pivots.active()) {
      this.c.shell.alert("Select a cell in a table or PivotTable to insert a slicer for it.")
      return
    }
    this.c.shell.openDialog?.("insertSlicers")
  }

  /** The columns the Insert Slicers dialog lists: the table's, or the PivotTable's fields (not calculated ones). */
  insertColumns(): string[] {
    const table = this.c.tables.active()
    if (table) return table.columns.map((c) => c.name)
    const pivot = this.c.pivots.activeWithCache()
    return pivot?.cache.fields.filter((f) => f.formula === undefined).map((f) => f.name) ?? []
  }

  /** A slicer name not taken in the workbook: the column's, then "Region 1", "Region 2"… */
  private freeName(base: string, taken: Set<string>): string {
    let name = base
    for (let n = 1; taken.has(name.toLowerCase()); n++) name = `${base} ${n}`
    taken.add(name.toLowerCase())
    return name
  }

  /** One slicer per column picked, in the middle of the window, each a little below and right of the last. */
  insert(columns: readonly string[]): SheetSlicer[] {
    const table = this.c.tables.active()
    const pivot = table ? null : this.c.pivots.activeWithCache()
    if ((!table && !pivot) || !columns.length) return []
    const taken = new Set(this.all().map((s) => s.slicer.name.toLowerCase()))
    let id = nextObjectId(this.c.wb.sheets)
    const start = this.c.objects.middle(SLICER_SIZE.w, SLICER_SIZE.h)
    const made = columns.map((column, i): SheetSlicer => {
      const box = { ...start, x: start.x + i * 24, y: start.y + i * 24 }
      return {
        kind: "slicer",
        id: id++,
        name: this.freeName(column, taken),
        anchor: anchorOf(this.c.sheet, box),
        placement: "fixed",
        table: table ? table.id : -1,
        column,
        ...(pivot
          ? {
              pivot: {
                cacheId: pivot.cache.id,
                field: column,
                tables: [{ sheet: this.c.sheet.name, name: pivot.table.name }],
              },
            }
          : {}),
        caption: column,
        header: true,
        style: DEFAULT_SLICER_STYLE,
        columns: 1,
        buttonHeight: SLICER_BUTTON_HEIGHT,
        sort: "ascending",
        customLists: true,
        crossFilter: "showItemsWithDataAtTop",
      }
    })
    if (!this.c.objects.commit([...this.c.objects.list, ...made], "Insert Slicer")) return []
    this.c.objects.select(made.map((o) => o.id))
    this.c.shell.showTab?.("TabSlicer")
    return made
  }

  /** A change to every selected slicer, as one undo step. */
  private change(label: string, edit: (o: SheetSlicer) => void): boolean {
    const ids = this.selected().map((o) => o.id)
    return this.c.objects.update(
      label,
      (o) => {
        if (!isChart(o) && o.kind === "slicer") edit(o)
      },
      ids,
    )
  }

  /** Slicer Caption (the tab's box). */
  setCaption(caption: string) {
    return this.change("Slicer Caption", (o) => (o.caption = caption))
  }

  styleIndex(): number {
    return SLICER_STYLES.indexOf(this.active()?.style as (typeof SLICER_STYLES)[number])
  }

  setStyle(style: string) {
    return this.change("Slicer Style", (o) => (o.style = style))
  }

  /** Buttons › Columns (1–20000 in Excel; Cells takes 1–100). */
  setColumns(n: number) {
    if (!Number.isFinite(n) || n < 1) return false
    return this.change("Slicer Columns", (o) => {
      o.columns = Math.min(100, Math.round(n))
      delete o.startItem
    })
  }

  /** Buttons › Height (px). */
  setButtonHeight(px: number) {
    if (!Number.isFinite(px) || px < 4) return false
    return this.change("Button Height", (o) => (o.buttonHeight = Math.round(px)))
  }

  /** A slicer's button width (px), from its width and columns. */
  buttonWidth(o: SheetSlicer): number {
    const w = this.c.objects.box(o).w
    const columns = Math.max(1, o.columns)
    return Math.max(4, (w - SIDES - GAP * (columns - 1)) / columns)
  }

  /** Buttons › Width: the slicer grows or shrinks to fit its buttons. */
  setButtonWidth(px: number) {
    if (!Number.isFinite(px) || px < 4) return false
    return this.change("Button Width", (o) => {
      const box = this.c.objects.box(o)
      const columns = Math.max(1, o.columns)
      const w = px * columns + GAP * (columns - 1) + SIDES
      o.anchor = anchorOf(this.c.sheet, { ...box, w })
    })
  }

  /** Slicer Settings. Returns Excel's message when the name is taken, else null. */
  applySettings(o: SheetSlicer, s: SlicerSettings): string | null {
    const name = s.name.trim()
    if (!name) return "Type a name for the slicer."
    const clash = this.all().some((x) => x.slicer.id !== o.id && x.slicer.name.toLowerCase() === name.toLowerCase())
    if (clash) return "The slicer name is already in use. Type a different name."
    this.c.objects.update(
      "Slicer Settings",
      (x) => {
        if (isChart(x) || x.kind !== "slicer") return
        Object.assign(x, { ...s, name })
        if (!s.hideNoData) delete x.hideNoData
      },
      [o.id],
    )
    return null
  }

  /** Every slicer of the workbook (with its sheet). */
  all(): { sheet: number; slicer: SheetSlicer }[] {
    return this.c.wb.sheets.flatMap((s, sheet) =>
      s.objects.flatMap((o) => (!isChart(o) && o.kind === "slicer" ? [{ sheet, slicer: o }] : [])),
    )
  }
}
