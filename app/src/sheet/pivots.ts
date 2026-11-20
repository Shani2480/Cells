import {
  chartRefRange,
  clearContentsOp,
  withCalcField,
  withItemsGrouped,
  VALUES_FIELD,
  withRangeGroup,
  withUngrouped,
  withFieldRemoved,
  groupKind,
  groupSettings,
  derivedDateGroups,
  type GroupKind,
  pivotChartButtons,
  pivotFromRecipe,
  recommendedPivots,
  type PivotRecipe,
  pivotChartData,
  type PivotChartLink,
  type SheetChart,
  type GroupSettings,
  fieldShared,
  withFieldIn,
  withoutCalcField,
  getPivotDataFormula,
  PIVOT_MESSAGE,
  currentRegion,
  newPivotTable,
  pivotIsEmpty,
  parseRange,
  pivotFieldName,
  pivotOptions,
  withAllCleared,
  withFiltersCleared,
  type Range,
  pivotSettingsTarget,
  withFieldSettings,
  withPivotOptions,
  withValueSettings,
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
  type PivotFieldSettings,
  type PivotOptions,
  type PivotSettingsTarget,
  type PivotShowAsSetting,
  type PivotValueSettings,
  type PivotCell,
  type PivotChoices,
  type PivotFilterSpec,
  type PivotFilterTarget,
  type PivotSort,
  type PivotTable,
  type Table,
} from "@cells/core"
import { withSheetData, type ChartSpec } from "@cells/chart"
import type { SheetController } from "./controller"

/** The width autofit adds for a cell's filter button (17 px and a gap). */
const FILTER_BUTTON_ROOM = 19

/** Where Generate GetPivotData's "off" is kept. */
const GETPIVOTDATA_KEY = "cells.generateGetPivotData"

/** Excel's message when a new PivotTable would sit over another. */
export const PIVOT_OVERLAP_MESSAGE = "A PivotTable report cannot overlap another PivotTable report."

/** Where Create PivotTable puts the new PivotTable. */
/** Recommended PivotTables with nothing to work from. */
export const RECOMMEND_NO_DATA =
  "Select a cell in a range of data with headings, or a table, to see recommended PivotTables."

export type PivotDestination = { kind: "newSheet" } | { kind: "existing"; sheet: number; at: CellAddress }

/** Which dialog a PivotTable filter dropdown opens, for the field at a filter button's cell. */
export type PivotFilterRequest = {
  kind: "label" | "value" | "top"
  /** The PivotTable's sheet when a PivotChart's button asked (it may not be the sheet shown). */
  sheet?: number
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
  refreshOps(
    cacheIds: readonly number[],
    sources?: ReadonlyMap<number, PivotCacheSource>,
    replace?: { cache: PivotCache; edit?: (t: PivotTable, sheet: number) => PivotTable },
  ): { ops: Op[]; overwrites: string[] } | string {
    const wb = this.c.wb
    const caches: PivotCache[] = [...wb.pivotCaches]
    const tables = new Map<number, PivotTable[]>(wb.sheets.map((s, i) => [i, [...s.pivots]]))
    const cellOps: Op[] = []
    const overwrites: string[] = []
    for (const id of cacheIds) {
      const index = caches.findIndex((c) => c.id === id)
      if (index < 0) continue
      const old = caches[index]
      let next: PivotCache
      if (replace?.cache.id === id) next = replace.cache
      else {
        const source = sources?.get(id) ?? old.source
        const data = readPivotSource(wb.sheets, wb.styles, source)
        if (typeof data === "string") return data
        next = { ...refreshedCache(old, data), source }
      }
      caches[index] = next
      wb.sheets.forEach((sheet, s) => {
        const list = tables.get(s)!
        list.forEach((t, k) => {
          if (t.cacheId !== id) return
          let laid = refreshedTable(t, old, next)
          if (replace?.edit) laid = replace.edit(laid, s)
          const plan = pivotCellsPlan(wb.sheets, wb.styles, s, t, laid, next)
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

  private run(
    cacheIds: number[],
    label: string,
    sources?: ReadonlyMap<number, PivotCacheSource>,
    replace?: Parameters<PivotCommands["refreshOps"]>[2],
  ): boolean {
    const result = this.refreshOps(cacheIds, sources, replace)
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

  /**
   * What grouping acts on: the field of the active cell's label (or header), how it groups, and the items of that
   * field the selection covers.
   */
  groupTarget(): { table: PivotTable; cache: PivotCache; field: number; kind: GroupKind; xs: number[] } | null {
    const at = this.activeWithCache()
    if (!at) return null
    const { row, col } = this.c.selection.active
    const cell = pivotCellAt(at.table, at.cache, row, col)
    const target = this.settingsTargetAt(row, col)
    const field = cell?.item?.field ?? (target?.kind === "field" ? target.field : -1)
    if (field < 0 || field === VALUES_FIELD) return null
    const xs = new Set<number>()
    for (const r of this.c.selection.ranges) {
      // Labels only: a selection bigger than the table is cut to it.
      const area = at.table.location
      for (let y = Math.max(r.r0, area.r0); y <= Math.min(r.r1, area.r1); y++)
        for (let x = Math.max(r.c0, area.c0); x <= Math.min(r.c1, area.c1); x++) {
          const item = pivotCellAt(at.table, at.cache, y, x)?.item
          if (item?.field === field) xs.add(item.x)
        }
    }
    return { ...at, field, kind: groupKind(at.cache, field), xs: [...xs] }
  }

  /** The field whose Grouping dialog the active cell opens (a date part field opens its date's), or -1. */
  groupDialogField(): number {
    const t = this.groupTarget()
    if (!t) return -1
    const g = t.cache.fields[t.field]?.group
    if (g?.kind === "range") return g.base
    return t.kind === "items" ? -1 : t.field
  }

  /** Group Selection: dates and numbers open the Grouping dialog; other items become one group. */
  groupSelection(): boolean {
    const t = this.groupTarget()
    if (!t) return false
    if (this.groupDialogField() >= 0) {
      this.c.shell.openDialog?.("pivotGroup")
      return true
    }
    if (!t.xs.length) return false
    const made = withItemsGrouped(t.cache, t.field, t.xs)
    if (typeof made === "string") {
      this.c.shell.alert(made)
      return true
    }
    const isNew = made.cache.fields.length > t.cache.fields.length
    const name = made.cache.fields[made.field].name
    const baseName = t.cache.fields[t.field].name
    const edit = (laid: PivotTable) => {
      if (!isNew) return laid
      // The new field goes just outside the one it groups, its items in label order (Group1, Plums).
      const find = (n: string) => made.cache.fields.findIndex((f) => f.name === n)
      const base = find(baseName)
      const p = find(name)
      const area = laid.rows.includes(base) ? "rows" : laid.cols.includes(base) ? "cols" : null
      if (!area) return laid
      let next = withFieldIn(laid, made.cache, p, area, laid[area].indexOf(base))
      const labels = fieldShared(made.cache, p)
      const order = labels.map((_, x) => x).sort((a, b) => String(labels[a]).localeCompare(String(labels[b])))
      const fields = [...next.fields]
      fields[p] = { ...fields[p], items: order.map((x) => ({ x })) }
      next = { ...next, fields }
      return next
    }
    this.run([t.cache.id], "Group", undefined, { cache: made.cache, edit })
    return true
  }

  /** The Grouping dialog's values for the field it opens on. */
  groupDialog(): (GroupSettings & { autoStart: number; autoEnd: number; dates: boolean; name: string }) | null {
    const f = this.groupDialogField()
    const at = this.activeWithCache()
    if (f < 0 || !at) return null
    return { ...groupSettings(at.cache, f), dates: groupKind(at.cache, f) === "dates", name: at.cache.fields[f].name }
  }

  /**
   * Grouping dialog's OK: numbers in ranges on their field; dates by the parts chosen, those fields taking the date
   * field's place (or its parts') on every PivotTable of the cache. Excel's message when it can't.
   */
  setGroup(settings: GroupSettings): string | null {
    const f = this.groupDialogField()
    const at = this.activeWithCache()
    if (f < 0 || !at) return null
    const next = withRangeGroup(at.cache, f, settings)
    if (typeof next === "string") return next
    const old = at.cache
    const names = (c: PivotCache, base: number) => [
      c.fields[base].name,
      ...derivedDateGroups(c, base).map((d) => c.fields[d.index].name),
    ]
    const edit =
      groupKind(old, f) === "dates"
        ? this.placeFields(
            names(old, f),
            () => {
              // The new parts, largest outermost.
              const parts = derivedDateGroups(next, f).map((d) => d.index)
              return parts.reverse()
            },
            next,
          )
        : undefined
    this.run([old.id], "Group", undefined, { cache: next, edit })
    return null
  }

  /**
   * An edit for `refreshOps`: in each PivotTable, the first of the fields named `replaced` that sat on an axis gives
   * its place to `fields()` (the others named leave).
   */
  private placeFields(replaced: string[], fields: () => number[], next: PivotCache) {
    const sheets = this.c.wb.sheets
    const lower = new Set(replaced.map((n) => n.toLowerCase()))
    return (laid: PivotTable, s: number) => {
      const before = sheets[s]?.pivots.find((t) => t.name === laid.name)
      const cache = before && this.c.wb.pivotCaches.find((c) => c.id === before.cacheId)
      if (!before || !cache) return laid
      const isReplaced = (i: number) => i !== VALUES_FIELD && lower.has(cache.fields[i]?.name.toLowerCase() ?? "")
      let area: "rows" | "cols" | "filters" | null = null
      let index = 0
      for (const a of ["rows", "cols"] as const) {
        const i = before[a].findIndex(isReplaced)
        if (i >= 0 && area === null) {
          area = a
          index = before[a].slice(0, i).filter((x) => !isReplaced(x)).length
        }
      }
      if (area === null && before.pages.some((p) => isReplaced(p.field))) {
        area = "filters"
        index = before.pages.findIndex((p) => isReplaced(p.field))
      }
      if (area === null) return laid
      let t = laid
      for (const n of replaced) {
        const i = next.fields.findIndex((k) => k.name.toLowerCase() === n.toLowerCase())
        if (i >= 0) t = withFieldRemoved(t, i)
      }
      fields().forEach((i, k) => (t = withFieldIn(t, next, i, area!, index + k)))
      return t
    }
  }

  /** Ungroup: number ranges, a date's parts, or the selected groups of a field made by Group Selection. */
  ungroup(): boolean {
    const t = this.groupTarget()
    if (!t) return false
    const g = t.cache.fields[t.field].group
    const xs = g?.kind === "discrete" ? t.xs : undefined
    const done = withUngrouped(t.cache, t.field, xs)
    if (!done) return false
    const old = t.cache
    let edit: ((laid: PivotTable, s: number) => PivotTable) | undefined
    if (done.removed) {
      const gone = old.fields.filter(
        (_, i) => i !== done.base && !done.cache.fields.some((k) => k.name === old.fields[i].name),
      )
      const baseName = old.fields[done.base].name
      edit = this.placeFields(
        gone.map((k) => k.name),
        () => {
          const i = done.cache.fields.findIndex((k) => k.name === baseName)
          return i >= 0 ? [i] : []
        },
        done.cache,
      )
      // The base already on the axis (Product under Product2) stays where it is.
      const inner = edit
      edit = (laid, s) => {
        const i = done.cache.fields.findIndex((k) => k.name === baseName)
        const onAxis = [...laid.rows, ...laid.cols].includes(i)
        if (!onAxis) return inner(laid, s)
        let next = laid
        for (const k of gone) {
          const j = done.cache.fields.findIndex((x) => x.name === k.name)
          if (j >= 0) next = withFieldRemoved(next, j)
        }
        return next
      }
    }
    this.run([old.id], "Ungroup", undefined, { cache: done.cache, edit })
    return true
  }

  /** Ungroup is there: the active field is grouped, or is a date with parts. */
  canUngroup(): boolean {
    const t = this.groupTarget()
    if (!t) return false
    return !!t.cache.fields[t.field].group || derivedDateGroups(t.cache, t.field).length > 0
  }

  /**
   * A PivotChart field button's dropdown: the PivotTable's sheet, a cell of it whose filter button covers the field,
   * and the field (null when the PivotTable is gone or shows no filter buttons).
   */
  chartButtonTarget(
    link: PivotChartLink,
    area: "filters" | "axis" | "legend",
    index: number,
  ): { sheet: number; row: number; col: number; target: PivotFilterTarget } | null {
    const at = this.byLink(link)
    if (!at) return null
    const buttons = pivotChartButtons(at.table, at.cache)
    const field = buttons[area][index]?.field
    if (field === undefined) return null
    for (const cell of pivotLayout(at.table, at.cache).cells) {
      const target = pivotFilterTarget(at.table, at.cache, cell)
      if (target?.fields.includes(field))
        return { sheet: at.sheet, row: cell.row, col: cell.col, target: { kind: target.kind, fields: [field] } }
    }
    return null
  }

  /** Insert › PivotChart in a PivotTable: a PivotChart of the active PivotTable (null when there is none). */
  insertChart(variant = "column") {
    const t = this.active()
    if (!t) return null
    return this.c.charts.insertPivot({ sheet: this.c.sheet.name, name: t.name }, variant)
  }

  /**
   * Create PivotChart: a new PivotTable of `source` at `dest`, and a PivotChart of it next to it (as Excel does, on
   * the PivotTable's sheet). Returns a problem to show, or null.
   */
  createWithChart(source: string, dest: PivotDestination, variant = "column"): string | null {
    return this.create(source, dest, () => void this.insertChart(variant))
  }

  /** A PivotChart's PivotTable and cache, by the sheet and name it links to (null when gone). */
  byLink(link: PivotChartLink): { table: PivotTable; cache: PivotCache; sheet: number } | null {
    const wb = this.c.wb
    const sheet = wb.sheets.findIndex((s) => s.name.toLowerCase() === link.sheet.toLowerCase())
    const table = wb.sheets[sheet]?.pivots.find((t) => t.name.toLowerCase() === link.name.toLowerCase())
    const cache = table && wb.pivotCaches.find((c) => c.id === table.cacheId)
    return table && cache ? { table, cache, sheet } : null
  }

  /** The active PivotTable's calculated fields (Calculated Field's Name list). */
  calcFields(): { name: string; formula: string }[] {
    const at = this.activeWithCache()
    return (at?.cache.fields ?? []).flatMap((f) =>
      f.formula !== undefined ? [{ name: f.name, formula: f.formula }] : [],
    )
  }

  /**
   * Calculated Field's Add / Modify: the field on the active PivotTable's cache, shared by every PivotTable on it; a
   * new one goes into the active PivotTable's values, as Excel adds it. Excel's message when it can't.
   */
  setCalcField(name: string, formula: string): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const next = withCalcField(at.cache, name, formula)
    if (typeof next === "string") return next
    const isNew = next.fields.length > at.cache.fields.length
    const activeSheet = this.c.sheetIndex
    const activeName = at.table.name
    const edit = (t: PivotTable, s: number) =>
      isNew && s === activeSheet && t.name === activeName ? withFieldIn(t, next, next.fields.length - 1, "values") : t
    this.run([at.cache.id], "Calculated Field", undefined, { cache: next, edit })
    return null
  }

  /** Calculated Field's Delete: the field leaves the cache and every PivotTable on it. */
  deleteCalcField(name: string) {
    const at = this.activeWithCache()
    if (!at) return
    this.run([at.cache.id], "Calculated Field", undefined, { cache: withoutCalcField(at.cache, name) })
  }

  /** Refresh: the active PivotTable's cache, and so every PivotTable that shares it. */
  refresh(): boolean {
    const t = this.active()
    return t ? this.run([t.cacheId], "Refresh") : false
  }

  /** Refresh the PivotTables of one cache (a PivotChart's, from PivotChart Analyze). */
  refreshCache(cacheId: number): boolean {
    return this.run([cacheId], "Refresh")
  }

  /** Refresh All: every cache some PivotTable uses. */
  refreshAll(): boolean {
    const used = new Set(this.c.wb.sheets.flatMap((s) => s.pivots.map((t) => t.cacheId)))
    return used.size ? this.run([...used], "Refresh All") : false
  }

  /** The PivotTable at a cell, its cache, and the engine's cell there. */
  cellAt(row: number, col: number): { table: PivotTable; cache: PivotCache; cell: PivotCell | undefined } | null {
    const table = pivotAt(this.c.wb.sheets[this.via ?? this.c.sheetIndex] ?? this.c.sheet, row, col)
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
   * Recommended PivotTables for the data Insert › PivotTable would start from (`sourceText`, else the PivotTable's
   * data, or the table or data at the active cell): each suggestion with the PivotTable it makes, for the previews. A message when there is no
   * data to suggest from.
   */
  recommendations(
    sourceText?: string,
  ): { source: string; items: { recipe: PivotRecipe; table: PivotTable; cache: PivotCache }[] } | string {
    // In a PivotTable: its own data, as Excel suggests from it.
    const text = sourceText ?? (this.active() ? this.sourceText() : this.defaultSource())
    const source = text ? this.parseSource(text) : null
    if (!source) return RECOMMEND_NO_DATA
    const wb = this.c.wb
    const data = readPivotSource(wb.sheets, wb.styles, source)
    if (typeof data === "string") return data
    if (!data.records.length) return RECOMMEND_NO_DATA
    const cache = pivotCacheFrom(nextPivotCacheId(wb.pivotCaches), source, data)
    const items = recommendedPivots(cache).map((recipe) => ({
      recipe,
      ...pivotFromRecipe(nextPivotName(wb.sheets), cache, { row: 2, col: 0 }, recipe),
    }))
    return { source: text, items }
  }

  /**
   * Create PivotTable's OK: a new cache from the source and an empty PivotTable, on a new sheet before this one (at
   * A3, as Excel puts it) or at a cell, shown and selected so the Fields pane opens on it. A message when it can't.
   */
  create(sourceText: string, dest: PivotDestination, then?: () => void, recipe?: PivotRecipe): string | null {
    const source = this.parseSource(sourceText)
    if (!source) return "Reference isn't valid."
    const wb = this.c.wb
    const data = readPivotSource(wb.sheets, wb.styles, source)
    if (typeof data === "string") return data
    if (!data.records.length && source.kind === "range") return "Reference isn't valid."
    const raw = pivotCacheFrom(nextPivotCacheId(wb.pivotCaches), source, data)
    const newSheet = dest.kind === "newSheet"
    const index = newSheet ? this.c.sheetIndex : dest.sheet
    const at = newSheet ? { row: 2, col: 0 } : dest.at
    // A recommended PivotTable comes with its fields placed (and its dates grouped in its cache).
    const made = recipe
      ? pivotFromRecipe(nextPivotName(wb.sheets), raw, at, recipe)
      : { table: newPivotTable(nextPivotName(wb.sheets), raw, at), cache: raw }
    const { table, cache } = made
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
      then?.()
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

  /**
   * Generate GetPivotData (Analyze › Options, Excel's "Use GetPivotData functions for PivotTable references"): a
   * formula pointing at a value cell gets a GETPIVOTDATA call instead of a reference. On unless turned off; kept on
   * this device.
   */
  get generateGetPivotData(): boolean {
    try {
      return localStorage.getItem(GETPIVOTDATA_KEY) !== "off"
    } catch {
      return true
    }
  }

  set generateGetPivotData(on: boolean) {
    try {
      if (on) localStorage.removeItem(GETPIVOTDATA_KEY)
      else localStorage.setItem(GETPIVOTDATA_KEY, "off")
    } catch {
      // Storage blocked: the option stays as it was.
    }
    this.c.changed()
  }

  /** The GETPIVOTDATA call for a value cell of a PivotTable on the active sheet, when the option is on. */
  pointedFormula(row: number, col: number): string | null {
    if (!this.generateGetPivotData) return null
    const at = this.cellAt(row, col)
    return at ? getPivotDataFormula(at.table, at.cache, row, col) : null
  }

  /** The window shows the phone layout (App sets it). */
  phone = false

  /** Whether the Fields pane shows now, in the window's layout. */
  fieldsShown(phone = this.phone): boolean {
    return this.fieldsOn(phone) && this.active() !== null
  }

  /** Field List is on: the pane shows whenever the active cell is in a PivotTable. */
  fieldsOn(phone = this.phone): boolean {
    return this.fieldsOpen && (!phone || this.fieldsAsked)
  }

  /** Analyze › Field List, and the cell menu's Show/Hide Field List. */
  toggleFields() {
    this.setFieldsOpen(!this.fieldsOn())
  }

  /** The active PivotTable and its cache, for the Fields pane. */
  activeWithCache(): { table: PivotTable; cache: PivotCache } | null {
    const table = this.active()
    const cache = table && this.c.wb.pivotCaches.find((x) => x.id === table.cacheId)
    return table && cache ? { table, cache } : null
  }

  /**
   * The Fields pane's Defer Layout Update: its changes wait in `pending` (for the PivotTable they started from) until
   * Update lays them out as one step. Changes from anywhere else apply at once, as in Excel.
   */
  deferred = false
  private pending: { base: PivotTable; next: PivotTable } | null = null

  /** The PivotTable as the pane shows it: with the changes waiting for Update. */
  paneTable(): { table: PivotTable; cache: PivotCache; waiting: boolean } | null {
    const at = this.activeWithCache()
    if (!at) return null
    const pending = this.pending?.base === at.table ? this.pending.next : null
    return { table: pending ?? at.table, cache: at.cache, waiting: pending !== null }
  }

  /** A change from the Fields pane: laid out now, or kept for Update while layout updates are deferred. */
  paneChange(make: (table: PivotTable, cache: PivotCache) => PivotTable): void {
    const at = this.paneTable()
    if (!at) return
    if (!this.deferred) return void this.change(make)
    const base = this.activeWithCache()!.table
    this.pending = { base, next: make(at.table, at.cache) }
    this.c.changed()
  }

  /** The pane's Update button: the waiting changes laid out as one undo step. */
  update(): boolean {
    const at = this.activeWithCache()
    const pending = this.pending?.base === at?.table ? this.pending : null
    this.pending = null
    if (!at || !pending) return false
    return this.relayout(at, pending.next, "PivotTable Field")
  }

  /** Defer Layout Update on or off (off lays out what is waiting). */
  setDeferred(on: boolean) {
    if (!on) this.update()
    this.deferred = on
    this.c.changed()
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

  // ── Value Field Settings, Field Settings, PivotTable Options ──

  /** The value field or field the cell menu's settings commands work on at a cell (null: none). */
  settingsTargetAt(row: number, col: number): PivotSettingsTarget | null {
    const table = pivotAt(this.c.sheet, row, col)
    const cache = table && this.c.wb.pivotCaches.find((x) => x.id === table.cacheId)
    return table && cache ? pivotSettingsTarget(table, cache, row, col) : null
  }

  /**
   * A double-click on a header opens its settings, as in Excel: a value field's name opens Value Field Settings, a
   * field's header or report filter label Field Settings. False elsewhere (items toggle, values show details).
   */
  openSettingsAt(row: number, col: number): boolean {
    const table = pivotAt(this.c.sheet, row, col)
    const cache = table && this.c.wb.pivotCaches.find((x) => x.id === table.cacheId)
    const cell = table && cache ? pivotCellAt(table, cache, row, col) : undefined
    if (!cell || cell.item || cell.lines) return false
    if (!["dataHeader", "corner", "rowHeader", "colHeader", "pageLabel", "colItem", "rowItem"].includes(cell.role))
      return false
    const t = this.settingsTargetAt(row, col)
    if (!t) return false
    this.c.shell.openDialog?.(t.kind === "value" ? `pivotValue:${t.index}` : `pivotField:${t.field}`)
    return true
  }

  /** Value Field Settings' OK for value field `index` of the active PivotTable; Excel's message when it can't. */
  setValueSettings(index: number, settings: PivotValueSettings): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const next = withValueSettings(at.table, at.cache, index, settings)
    if (typeof next === "string") return next
    this.relayout(at, next, "Value Field Settings")
    return null
  }

  /** Field Settings' OK for field `f` of the active PivotTable. */
  setFieldSettings(f: number, settings: PivotFieldSettings): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const next = withFieldSettings(at.table, at.cache, f, settings)
    if (typeof next === "string") return next
    this.relayout(at, next, "Field Settings")
    return null
  }

  /** PivotTable Options' OK for the active PivotTable (Refresh data when opening the file goes to its cache). */
  setOptions(options: PivotOptions): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const next = withPivotOptions(at.table, this.c.sheet, options)
    if (typeof next === "string") return next
    const caches = this.c.wb.pivotCaches.map((x) =>
      x === at.cache && x.refreshOnLoad !== options.refreshOnLoad ? { ...x, refreshOnLoad: options.refreshOnLoad } : x,
    )
    const extra: Op[] = caches.some((x, i) => x !== this.c.wb.pivotCaches[i]) ? [{ kind: "pivotCaches", caches }] : []
    this.relayout(at, next, "PivotTable Options", extra)
    return null
  }

  // ── The PivotTable Analyze and Design tabs ──

  /** Analyze › PivotTable Name. */
  rename(name: string): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const next = withPivotOptions(at.table, this.c.sheet, { ...pivotOptions(at.table, at.cache), name })
    if (typeof next === "string") return next
    this.relayout(at, next, "Rename PivotTable")
    return null
  }

  /** The Active Field box: the value field's or field's name at the active cell (null: none there). */
  activeField(): { target: PivotSettingsTarget; name: string } | null {
    const at = this.activeWithCache()
    const { row, col } = this.c.selection.active
    const target = at && this.settingsTargetAt(row, col)
    if (!at || !target) return null
    const name =
      target.kind === "value"
        ? (at.table.values[target.index]?.name ?? "")
        : pivotFieldName(at.table, at.cache, target.field)
    return { target, name }
  }

  /** The Active Field box's new name, for the value field or field at the active cell. */
  renameActiveField(name: string): string | null {
    const at = this.activeWithCache()
    const active = this.activeField()
    if (!at || !active || name === active.name) return null
    const { target } = active
    if (target.kind === "value") {
      const v = at.table.values[target.index]
      return this.setValueSettings(target.index, { name, fn: v.fn, numFmt: v.numFmt })
    }
    const f = at.table.fields[target.field]
    return this.setFieldSettings(target.field, { ...f, name })
  }

  /** Analyze › Change Data Source: the active PivotTable's cache reads another table or range. */
  changeSource(text: string): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const source = this.parseSource(text)
    if (!source) return "Reference isn't valid."
    const data = readPivotSource(this.c.wb.sheets, this.c.wb.styles, source)
    if (typeof data === "string") return data
    this.run([at.cache.id], "Change Data Source", new Map([[at.cache.id, source]]))
    return null
  }

  /** The active PivotTable's source as the Table/Range box shows it. */
  sourceText(): string {
    const source = this.activeWithCache()?.cache.source
    if (!source) return ""
    if (source.kind === "name") return source.name
    return `${quoteSheet(source.sheet)}!${absolute(parseRange(source.ref) ?? { r0: 0, c0: 0, r1: 0, c1: 0 })}`
  }

  /** Analyze › Clear › Clear All and Clear Filters. */
  clear(what: "all" | "filters"): boolean {
    return this.change(
      (t, cache) => (what === "all" ? withAllCleared(t, cache) : withFiltersCleared(t)),
      what === "all" ? "Clear All" : "Clear Filters",
    )
  }

  /**
   * Analyze › Select: the entire PivotTable (report filters too), its labels and values (the table), its values, or
   * its labels.
   */
  select(what: "entire" | "labelsAndValues" | "values" | "labels"): boolean {
    const at = this.activeWithCache()
    if (!at) return false
    const { table, cache } = at
    const layout = pivotLayout(table, cache)
    const r = layout.range
    const data = { r0: r.r0 + layout.firstDataRow, c0: r.c0 + layout.firstDataCol, r1: r.r1, c1: r.c1 }
    let ranges: Range[]
    if (what === "entire") ranges = [pivotArea(table)]
    else if (what === "labelsAndValues") ranges = [r]
    else if (what === "values") ranges = [data]
    else
      ranges = [
        ...(data.c0 > r.c0 ? [{ r0: r.r0, c0: r.c0, r1: r.r1, c1: data.c0 - 1 }] : []),
        ...(data.r0 > r.r0 ? [{ r0: r.r0, c0: data.c0, r1: data.r0 - 1, c1: r.c1 }] : []),
      ]
    ranges = ranges.filter((x) => x.r1 >= x.r0 && x.c1 >= x.c0)
    if (!ranges.length) return false
    const first = ranges[0]
    const corner = { row: first.r0, col: first.c0 }
    const far = { row: first.r1, col: first.c1 }
    this.c.setSelection({ ranges, active: corner, activeRange: 0, anchor: corner, focus: far }, null)
    return true
  }

  /**
   * Analyze › Move PivotTable: to a cell of this or another sheet (its report filters' top left), or a new sheet.
   * A message when it can't.
   */
  move(dest: PivotDestination): string | null {
    const at = this.activeWithCache()
    if (!at) return null
    const wb = this.c.wb
    const from = this.c.sheetIndex
    const newSheet = dest.kind === "newSheet"
    const sheets = [...wb.sheets]
    const name = nextSheetName(wb.sheets.map((s) => s.name))
    const to = newSheet ? from : dest.sheet
    if (newSheet) sheets.splice(to, 0, new Sheet(name))
    const fromAfter = newSheet ? from + 1 : from
    const top = newSheet ? { row: 2, col: 0 } : dest.at
    const pagesRows = at.table.pages.length ? at.table.pages.length + 1 : 0
    const moved = { ...at.table, at: { row: top.row + pagesRows, col: top.col } }
    const before = to === fromAfter ? at.table : null
    const others = sheets[to].pivots.filter((t) => t !== at.table)
    const area = pivotArea({ ...moved, location: pivotLayout(moved, at.cache).range })
    const hits = (r: Range) => r.r0 <= area.r1 && r.r1 >= area.r0 && r.c0 <= area.c1 && r.c1 >= area.c0
    if (others.some((t) => hits(pivotArea(t)))) return PIVOT_OVERLAP_MESSAGE
    const plan = pivotCellsPlan(sheets, wb.styles, to, before, moved, at.cache)
    const ops: Op[] = newSheet ? [{ kind: "addSheet", sheet: to, name }] : []
    if (!before) {
      // Off its old sheet: its cells there are cleared and it leaves that sheet's list.
      const old = pivotArea(at.table)
      const cells = []
      for (let r = old.r0; r <= old.r1; r++)
        for (let c = old.c0; c <= old.c1; c++)
          if (wb.sheets[from].cells.get(r, c)) cells.push({ row: r, col: c, data: undefined })
      ops.push({ kind: "setCells", sheet: fromAfter, cells })
      ops.push({ kind: "pivots", sheet: fromAfter, tables: wb.sheets[from].pivots.filter((t) => t !== at.table) })
    }
    ops.push(...plan.ops, { kind: "pivots", sheet: to, tables: [...others, plan.table] })
    if (to !== fromAfter)
      ops.push(...pivotChartLinkOps(sheets, wb.sheets[from].name, at.table.name, sheets[to].name, at.table.name))
    const go = () => {
      const entry = this.c.transact(ops, { label: "Move PivotTable", before: this.c.selection, sheetAfter: to })
      if (!entry) return
      if (to !== this.c.sheetIndex) this.c.activate(to)
      this.fitColumns(entry, ops)
      this.c.goTo({ r0: plan.table.at.row, c0: plan.table.at.col, r1: plan.table.at.row, c1: plan.table.at.col })
    }
    if (plan.overwrites) this.c.shell.confirm(pivotOverwriteMessage(this.c.printFile.fileName, sheets[to].name), go)
    else go()
    return null
  }

  /**
   * Delete over a selection that holds a whole PivotTable (its report filters too) deletes it, as in Excel, with the
   * rest of the selection cleared and its cache dropped when nothing else uses it. False when no PivotTable is wholly
   * selected; a PivotTable only partly selected still gets Excel's message.
   */
  deleteCovered(ranges: readonly Range[]): boolean {
    const sheet = this.c.sheet
    const inside = (a: Range) => ranges.some((r) => a.r0 >= r.r0 && a.r1 <= r.r1 && a.c0 >= r.c0 && a.c1 <= r.c1)
    const touches = (a: Range) => ranges.some((r) => a.r0 <= r.r1 && a.r1 >= r.r0 && a.c0 <= r.c1 && a.c1 >= r.c0)
    const gone = sheet.pivots.filter((t) => inside(pivotArea(t)))
    if (!gone.length) return false
    if (sheet.pivots.some((t) => !gone.includes(t) && touches(pivotArea(t)))) {
      this.c.shell.alert(PIVOT_MESSAGE)
      return true
    }
    const s = this.c.sheetIndex
    const wb = this.c.wb
    const left = sheet.pivots.filter((t) => !gone.includes(t))
    const used = new Set(wb.sheets.flatMap((x) => (x === sheet ? left : x.pivots).map((t) => t.cacheId)))
    const caches = wb.pivotCaches.filter((c) => used.has(c.id))
    const clear = clearContentsOp(s, sheet, ranges)
    const ops: Op[] = [
      { kind: "pivots", sheet: s, tables: left },
      ...(clear ? [clear] : []),
      ...(caches.length !== wb.pivotCaches.length ? [{ kind: "pivotCaches" as const, caches }] : []),
    ]
    ops.push(
      ...pivotChartOps(
        wb.sheets,
        sheet.name,
        gone.map((t) => t.name),
        (c) => {
          const t = gone.find((x) => x.name.toLowerCase() === c.pivot!.name.toLowerCase())!
          const cache = wb.pivotCaches.find((x) => x.id === t.cacheId)
          return cache ? staticPivotChart(c, t, cache) : c
        },
      ),
    )
    this.c.transact(ops, { label: "Clear", before: this.c.selection, after: this.c.selection })
    return true
  }

  /** Design › Layout and PivotTable Style Options, and the style gallery. */
  layout(make: (t: PivotTable) => PivotTable, label: string): boolean {
    return this.change((t) => make(t), label)
  }

  setStyle(name: string | null): boolean {
    return this.change((t) => ({ ...t, style: { ...t.style, name } }), "PivotTable Style")
  }

  setStyleOption(key: "rowHeaders" | "colHeaders" | "rowStripes" | "colStripes", on: boolean): boolean {
    return this.change((t) => ({ ...t, style: { ...t.style, [key]: on } }), "PivotTable Style Options")
  }

  /** Analyze › Show › +/- Buttons. */
  setButtons(on: boolean): boolean {
    return this.change((t) => {
      const next = { ...t }
      if (on) delete next.showDrill
      else next.showDrill = false
      return next
    }, "+/- Buttons")
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
    // The source's columns only: calculated fields have no records (Excel leaves them out too).
    const fields = at.cache.fields.filter((f) => f.formula === undefined)
    const styleOf = fields.map((f) =>
      f.numFmt && f.numFmt !== "General" ? wb.styles.intern({ ...wb.styles.get(0), numberFormat: f.numFmt }) : 0,
    )
    const cells: CellWrite[] = fields.map((f, col) => ({ row: 0, col, data: { value: f.name } }))
    records.forEach((r, i) =>
      at.cache.records[r].forEach((value, col) => {
        if (value === null || value === "" || col >= fields.length) return
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
      at.table.fields[field].items.length <
        (at.cache.fields[field].items || at.cache.fields[field].group ? fieldShared(at.cache, field).length : 0) ||
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

  /** Show Values As' Base Item list for field `f` of the active PivotTable: its items as shown, labels formatted. */
  baseItems(f: number): { x: number; label: string }[] {
    const at = this.activeWithCache()
    if (!at || !at.table.fields[f]) return []
    return pivotChoices(at.table, at.cache, f, (v, numFmt) => displayText(v, numFmt ?? "General")).items.map(
      ({ x, label }) => ({ x, label }),
    )
  }

  /** The cell menu's Show Values As for the value field of the active cell (one undo step). */
  setShowAs(index: number, show: PivotShowAsSetting): string | null {
    const v = this.activeWithCache()?.table.values[index]
    if (!v) return null
    return this.setValueSettings(index, { name: v.name, fn: v.fn, numFmt: v.numFmt, show })
  }

  /**
   * The sheet the filter dropdown works on while a PivotChart's button drives it (its PivotTable may be on another
   * sheet); null: the sheet shown.
   */
  private via: number | null = null

  /**
   * The filter dropdown's commands for the PivotTables of sheet `sheet` (a PivotChart's buttons: its PivotTable can be
   * on another sheet than the chart); the sheet shown when undefined.
   */
  on(sheet?: number): PivotFilterCommands {
    if (sheet === undefined || sheet === this.c.sheetIndex) return this
    const at = <T>(run: () => T): T => {
      const was = this.via
      this.via = sheet
      try {
        return run()
      } finally {
        this.via = was
      }
    }
    return {
      choices: (r, c, f) => at(() => this.choices(r, c, f)),
      showItems: (r, c, f, shown) => at(() => this.showItems(r, c, f, shown)),
      setFilter: (r, c, f, spec) => at(() => this.setFilter(r, c, f, spec)),
      clearField: (r, c, f) => at(() => this.clearField(r, c, f)),
      valueNames: (r, c) => at(() => this.valueNames(r, c)),
      pickPage: (r, c, f, x) => at(() => this.pickPage(r, c, f, x)),
      sortField: (r, c, f, sort) => at(() => this.sortField(r, c, f, sort)),
    }
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

  private relayout(
    at: { table: PivotTable; cache: PivotCache },
    next: PivotTable,
    label: string,
    extra: Op[] = [],
  ): boolean {
    if (next === at.table && !extra.length) return false
    // From a PivotChart's button the PivotTable may be on another sheet: its cells change, the selection stays.
    const away = this.via !== null && this.via !== this.c.sheetIndex
    const s = away ? this.via! : this.c.sheetIndex
    const sheet = this.c.wb.sheets[s]
    const plan = pivotCellsPlan(this.c.wb.sheets, this.c.wb.styles, s, at.table, next, at.cache)
    const tables = sheet.pivots.map((t) => (t === at.table ? plan.table : t))
    const ops: Op[] = [...extra, ...plan.ops, { kind: "pivots", sheet: s, tables }]
    if (next.name !== at.table.name)
      ops.push(...pivotChartLinkOps(this.c.wb.sheets, sheet.name, at.table.name, sheet.name, next.name))
    const go = () => {
      const entry = this.c.transact(ops, { label, before: this.c.selection })
      if (entry) this.fitColumns(entry, ops)
      if (away) return
      // The active cell stays in the PivotTable (at its top left) when the change shrinks the table away from it, so
      // the Fields pane stays open on it.
      const { row, col } = this.c.selection.active
      const area = pivotArea(plan.table)
      if (row < area.r0 || row > area.r1 || col < area.c0 || col > area.c1)
        this.c.goTo({ r0: plan.table.at.row, c0: plan.table.at.col, r1: plan.table.at.row, c1: plan.table.at.col })
    }
    if (plan.overwrites) this.c.shell.confirm(pivotOverwriteMessage(this.c.printFile.fileName, sheet.name), go)
    else go()
    return true
  }
}

/** What a PivotTable filter dropdown asks of the PivotTables (PivotCommands.on gives them for another sheet). */
export type PivotFilterCommands = Pick<
  PivotCommands,
  "choices" | "showItems" | "setFilter" | "clearField" | "valueNames" | "pickPage" | "sortField"
>

/** A range with $ signs, as Excel's Table/Range box shows it. */
function absolute(r: { r0: number; c0: number; r1: number; c1: number }): string {
  return rangeName(r)
    .split(":")
    .map((part) => part.replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2"))
    .join(":")
}

/**
 * The ops that point every PivotChart of PivotTable `name` on sheet `sheet` at its new sheet and name (`sheets` as
 * the ops number them).
 */
function pivotChartLinkOps(
  sheets: readonly Sheet[],
  sheet: string,
  name: string,
  toSheet: string,
  toName: string,
): Op[] {
  return pivotChartOps(sheets, sheet, [name], (c) => ({ ...c, pivot: { ...c.pivot!, sheet: toSheet, name: toName } }))
}

/** The ops that change every PivotChart of the PivotTables `names` on sheet `sheet` with `make`. */
function pivotChartOps(
  sheets: readonly Sheet[],
  sheet: string,
  names: readonly string[],
  make: (c: SheetChart) => SheetChart,
): Op[] {
  const ops: Op[] = []
  const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase()
  const hit = (c: SheetChart) => !!c.pivot && same(c.pivot.sheet, sheet) && names.some((n) => same(c.pivot!.name, n))
  sheets.forEach((s, i) => {
    if (!s.charts.some(hit)) return
    ops.push({ kind: "charts", sheet: i, charts: s.charts.map((c) => (hit(c) ? make(c) : c)) })
  })
  return ops
}

/** A PivotChart whose PivotTable is deleted keeps the points it showed, as a chart of its own (as in Excel). */
function staticPivotChart(chart: SheetChart, table: PivotTable, cache: PivotCache): SheetChart {
  const data = pivotChartData(table, cache, chart.pivot!.sheet)
  const { pivot: _, ...rest } = chart
  const spec = withSheetData(chart.spec as Partial<ChartSpec>, data)
  if (data.levels.length) spec.levels = data.levels
  return { ...rest, refs: null, spec: spec as unknown as Record<string, unknown> }
}
