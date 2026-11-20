import type { Sheet } from "../sheet"
import { pivotFieldName, pivotLayout } from "./engine"
import { pivotFilterTarget } from "./filter"
import {
  newPivotField,
  PIVOT_FN_LABEL,
  type PivotCache,
  type PivotField,
  type PivotFn,
  type PivotSubtotal,
  type PivotTable,
} from "./model"

/**
 * Value Field Settings, Field Settings and PivotTable Options: a PivotTable with their changes, or Excel's message
 * when a name can't be used. Laying the cells out again is the caller's.
 */

/** What the cell menu's settings commands work on at a PivotTable cell: a value field, or a row, column or page field. */
export type PivotSettingsTarget = { kind: "value"; index: number } | { kind: "field"; field: number }

/** The value field or field a PivotTable cell belongs to (null: none, e.g. an empty corner). */
export function pivotSettingsTarget(
  table: PivotTable,
  cache: PivotCache,
  row: number,
  col: number,
): PivotSettingsTarget | null {
  const layout = pivotLayout(table, cache)
  const cell = layout.cells.find((c) => c.row === row && c.col === col)
  if (!cell) return null
  const value = (i: number | undefined): PivotSettingsTarget | null =>
    table.values.length && i !== undefined ? { kind: "value", index: Math.min(i, table.values.length - 1) } : null
  if (cell.lines) {
    const [ri, ci] = cell.lines
    return value((table.rows.includes(-2) ? layout.rowItems[ri]?.i : layout.colItems[ci]?.i) ?? 0)
  }
  if (cell.role === "dataHeader" || (cell.role === "corner" && cell.value !== undefined && table.values.length === 1))
    return value(0)
  if (cell.item) return { kind: "field", field: cell.item.field }
  // A value field's name among the column or row labels.
  if (cell.role === "colItem") return value(layout.colItems[col - layout.range.c0 - layout.firstDataCol]?.i)
  if (cell.role === "rowItem") return value(layout.rowItems[row - layout.range.r0 - layout.firstDataRow]?.i)
  const target = pivotFilterTarget(table, cache, cell)
  const field = target?.fields.find((f) => f >= 0)
  return field === undefined ? null : { kind: "field", field }
}

export const PIVOT_FIELD_NAME_TAKEN = "PivotTable field name already exists."
export const PIVOT_NAME_INVALID = "PivotTable name is not valid."

/** The functions in Value Field Settings' "Summarize value field by" list, in Excel's order. */
export const PIVOT_FNS: PivotFn[] = [
  "sum",
  "count",
  "average",
  "max",
  "min",
  "product",
  "countNums",
  "stdDev",
  "stdDevp",
  "var",
  "varp",
]

/** A value field's name as Excel makes it: "Sum of Units". */
export const autoDataFieldName = (table: PivotTable, cache: PivotCache, field: number, fn: PivotFn) =>
  `${PIVOT_FN_LABEL[fn]} of ${pivotFieldName(table, cache, field)}`

/** What Value Field Settings changes. */
export type PivotValueSettings = { name: string; fn: PivotFn; numFmt?: string }

/**
 * Value field `index` with new settings. A name is unique among the value fields and differs from every source
 * field's name (ignoring case; "Units " with a space is Excel's way round that). An empty name, or the automatic name
 * kept while the function changes, becomes the new function's automatic name.
 */
export function withValueSettings(
  table: PivotTable,
  cache: PivotCache,
  index: number,
  settings: PivotValueSettings,
): PivotTable | string {
  const v = table.values[index]
  if (!v) return table
  let name = settings.name
  const wasAuto = name === v.name && v.name === autoDataFieldName(table, cache, v.field, v.fn)
  if (!name || (wasAuto && settings.fn !== v.fn)) name = autoDataFieldName(table, cache, v.field, settings.fn)
  const lower = name.toLowerCase()
  const taken =
    table.values.some((o, i) => i !== index && o.name.toLowerCase() === lower) ||
    cache.fields.some((f, i) => f.name.toLowerCase() === lower || table.fields[i]?.name?.toLowerCase() === lower)
  if (taken) return PIVOT_FIELD_NAME_TAKEN
  const next = { ...v, name, fn: settings.fn }
  if (settings.numFmt && settings.numFmt !== "General") next.numFmt = settings.numFmt
  else delete next.numFmt
  const values = table.values.map((o, i) => (i === index ? next : o))
  return { ...table, values }
}

/** What Field Settings changes for a row, column or report filter field. */
export type PivotFieldSettings = {
  /** The custom name; empty or the source's own for none. */
  name: string
  subtotals: PivotSubtotal[]
  outline: boolean
  compact: boolean
  subtotalTop: boolean
  repeatLabels: boolean
  insertBlankRow: boolean
  showAll: boolean
}

/** Field `f` with new settings; its name must differ from the other fields' and the value fields' names. */
export function withFieldSettings(
  table: PivotTable,
  cache: PivotCache,
  f: number,
  settings: PivotFieldSettings,
): PivotTable | string {
  const field = table.fields[f]
  if (!field) return table
  const own = cache.fields[f]?.name ?? ""
  const name = settings.name && settings.name !== own ? settings.name : undefined
  if (name) {
    const lower = name.toLowerCase()
    const taken =
      table.values.some((v) => v.name.toLowerCase() === lower) ||
      cache.fields.some((_, i) => i !== f && pivotFieldName(table, cache, i).toLowerCase() === lower)
    if (taken) return PIVOT_FIELD_NAME_TAKEN
  }
  const next: PivotField = {
    ...field,
    subtotals: [...settings.subtotals],
    outline: settings.outline,
    compact: settings.compact,
    subtotalTop: settings.subtotalTop,
    repeatLabels: settings.repeatLabels,
    insertBlankRow: settings.insertBlankRow,
    showAll: settings.showAll,
  }
  if (name) next.name = name
  else delete next.name
  return { ...table, fields: table.fields.map((x, i) => (i === f ? next : x)) }
}

/** What PivotTable Options changes. */
export type PivotOptions = {
  name: string
  /** Layout & Format. */
  showError: boolean
  errorCaption: string
  showMissing: boolean
  missingCaption: string
  autoFit: boolean
  /** Totals & Filters. */
  rowGrandTotals: boolean
  colGrandTotals: boolean
  /** Display. */
  showDrill: boolean
  valuesRow: boolean
  /** Data. */
  enableDrill: boolean
  refreshOnLoad: boolean
}

/** The options a PivotTable has now, for the dialog. */
export function pivotOptions(table: PivotTable, cache: PivotCache): PivotOptions {
  return {
    name: table.name,
    showError: table.showError,
    errorCaption: table.captions.error ?? "",
    showMissing: table.showMissing,
    missingCaption: table.captions.missing ?? "",
    autoFit: table.autoFit !== false,
    rowGrandTotals: table.rowGrandTotals,
    colGrandTotals: table.colGrandTotals,
    showDrill: table.showDrill !== false,
    valuesRow: !table.hideValuesRow,
    enableDrill: table.enableDrill !== false,
    refreshOnLoad: cache.refreshOnLoad,
  }
}

/**
 * The table with new options (refreshOnLoad belongs to the cache: the caller's). A name can't be empty; two
 * PivotTables on one sheet can't share one (other sheets may, as in Excel).
 */
export function withPivotOptions(table: PivotTable, sheet: Sheet, options: PivotOptions): PivotTable | string {
  const name = options.name.trim()
  if (!name) return PIVOT_NAME_INVALID
  if (sheet.pivots.some((t) => t !== table && t.name.toLowerCase() === name.toLowerCase()))
    return "A PivotTable report with that name already exists on this sheet."
  const captions = { ...table.captions }
  if (options.errorCaption) captions.error = options.errorCaption
  else delete captions.error
  if (options.missingCaption) captions.missing = options.missingCaption
  else delete captions.missing
  const next: PivotTable = {
    ...table,
    name,
    captions,
    showError: options.showError,
    showMissing: options.showMissing,
    rowGrandTotals: options.rowGrandTotals,
    colGrandTotals: options.colGrandTotals,
  }
  const flag = <K extends "autoFit" | "showDrill" | "enableDrill">(key: K, on: boolean) => {
    if (on) delete next[key]
    else next[key] = false as PivotTable[K]
  }
  flag("autoFit", options.autoFit)
  flag("showDrill", options.showDrill)
  flag("enableDrill", options.enableDrill)
  if (options.valuesRow) delete next.hideValuesRow
  else next.hideValuesRow = true
  return next
}

// ── The Design tab's Layout group, and Analyze › Clear ──

/** The row and column fields (not Σ Values). */
const axisFields = (t: PivotTable) => [...t.rows, ...t.cols].filter((f) => f >= 0)

/** Every field (not only those in use), as Excel's layout commands change them all. */
function withEveryField(t: PivotTable, patch: (f: PivotField) => Partial<PivotField>): PivotTable {
  return { ...t, fields: t.fields.map((f) => ({ ...f, ...patch(f) })) }
}

/** Design › Subtotals: Do Not Show, Show all Subtotals at Bottom / at Top of Group. */
export function withSubtotalsShown(t: PivotTable, place: "none" | "bottom" | "top"): PivotTable {
  return withEveryField(t, (f) =>
    place === "none"
      ? { subtotals: [] }
      : { subtotals: f.subtotals.length ? f.subtotals : ["default"], subtotalTop: place === "top" },
  )
}

/** Which Design › Subtotals choice the table shows now (null: mixed). */
export function pivotSubtotalsShown(t: PivotTable): "none" | "bottom" | "top" | null {
  const used = axisFields(t).map((f) => t.fields[f])
  if (!used.length) return null
  if (used.every((f) => !f.subtotals.length)) return "none"
  if (used.every((f) => f.subtotals.length && f.subtotalTop)) return "top"
  if (used.every((f) => f.subtotals.length && !f.subtotalTop)) return "bottom"
  return null
}

/** Design › Grand Totals ("for rows" is the Grand Total column, Excel's colGrandTotals). */
export function withGrandTotals(t: PivotTable, rows: boolean, cols: boolean): PivotTable {
  return { ...t, colGrandTotals: rows, rowGrandTotals: cols }
}

export type PivotReportLayout = "compact" | "outline" | "tabular"

/** Design › Report Layout: compact, outline or tabular form, for the table and every field. */
export function withReportLayout(t: PivotTable, layout: PivotReportLayout): PivotTable {
  const compact = layout === "compact"
  const outline = layout !== "tabular"
  return { ...withEveryField(t, () => ({ compact, outline })), compact, outline, compactData: compact }
}

/** The report layout the table shows now (null: its fields differ). */
export function pivotReportLayout(t: PivotTable): PivotReportLayout | null {
  const used = t.rows.filter((f) => f >= 0).map((f) => t.fields[f])
  const of = (f: { compact: boolean; outline: boolean }): PivotReportLayout =>
    !f.outline ? "tabular" : f.compact ? "compact" : "outline"
  if (!used.length) return of(t)
  const first = of(used[0])
  return used.every((f) => of(f) === first) ? first : null
}

/** Design › Report Layout › Repeat All Item Labels / Do Not Repeat Item Labels. */
export const withRepeatAllLabels = (t: PivotTable, on: boolean) => withEveryField(t, () => ({ repeatLabels: on }))

/** Design › Blank Rows › Insert Blank Line after Each Item / Remove Blank Line after Each Item. */
export const withBlankLines = (t: PivotTable, on: boolean) => withEveryField(t, () => ({ insertBlankRow: on }))

/** Analyze › Clear › Clear Filters: every item ticked, no label or value filters, report filters back to (All). */
export function withFiltersCleared(t: PivotTable): PivotTable {
  const next: PivotTable = {
    ...withEveryField(t, (f) => ({ items: f.items.map(({ hidden: _hidden, ...i }) => i) })),
    pages: t.pages.map((p) => ({ ...p, item: null })),
  }
  delete next.filters
  return next
}

/** Analyze › Clear › Clear All: no fields, filters, sorts or names left; the name, place, style and options stay. */
export function withAllCleared(t: PivotTable, cache: PivotCache): PivotTable {
  const next: PivotTable = {
    ...t,
    fields: cache.fields.map((f) => newPivotField(f.items?.length ?? 0, t)),
    rows: [],
    cols: [],
    pages: [],
    values: [],
  }
  delete next.filters
  return next
}
