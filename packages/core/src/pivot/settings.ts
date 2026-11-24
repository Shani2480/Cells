import type { Sheet } from "../sheet"
import { pivotFieldName, pivotLayout, type PivotAxisItem } from "./engine"
import { fullItems, pivotFilterTarget } from "./filter"
import { fieldShared } from "./group"
import { showAsNumFmt } from "./show-as"
import {
  newPivotField,
  PIVOT_FN_LABEL,
  SHOW_AS_BASE_FIELD,
  SHOW_AS_BASE_ITEM,
  VALUES_FIELD,
  type PivotCache,
  type PivotDataField,
  type PivotField,
  type PivotShowAs,
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

/** What Value Field Settings changes; `show` left out keeps Show Values As as it is. */
export type PivotValueSettings = { name: string; fn: PivotFn; numFmt?: string; show?: PivotShowAsSetting }

/** Show Values As with its base field and item (`as` null: No Calculation). */
export type PivotShowAsSetting = {
  as: PivotShowAs | null
  baseField?: number
  baseItem?: number | "previous" | "next"
}

/** Show Values As in Excel's order, with Excel's names. */
export const SHOW_AS_CHOICES: readonly { as: PivotShowAs | null; label: string }[] = [
  { as: null, label: "No Calculation" },
  { as: "percentOfTotal", label: "% of Grand Total" },
  { as: "percentOfCol", label: "% of Column Total" },
  { as: "percentOfRow", label: "% of Row Total" },
  { as: "percent", label: "% Of" },
  { as: "percentOfParentRow", label: "% of Parent Row Total" },
  { as: "percentOfParentCol", label: "% of Parent Column Total" },
  { as: "percentOfParent", label: "% of Parent Total" },
  { as: "difference", label: "Difference From" },
  { as: "percentDiff", label: "% Difference From" },
  { as: "runTotal", label: "Running Total In" },
  { as: "percentOfRunningTotal", label: "% Running Total In" },
  { as: "rankAscending", label: "Rank Smallest to Largest" },
  { as: "rankDescending", label: "Rank Largest to Smallest" },
  { as: "index", label: "Index" },
]

/** The base field Show Values As starts on: the first row field, else the first column field. */
export function defaultBaseField(table: PivotTable): number {
  return [...table.rows, ...table.cols].find((f) => f !== VALUES_FIELD) ?? 0
}

/** A value field with Show Values As set; a percentage gets Excel's 0.00% unless it has a format of its own. */
function withShowAs(v: PivotDataField, show: PivotShowAsSetting): PivotDataField {
  const next: PivotDataField = { ...v }
  delete next.showAs
  delete next.baseField
  delete next.baseItem
  const before = showAsNumFmt(v.showAs)
  if (next.numFmt === undefined || next.numFmt === before) {
    const after = show.as ? showAsNumFmt(show.as) : undefined
    if (after) next.numFmt = after
    else delete next.numFmt
  }
  if (!show.as) return next
  next.showAs = show.as
  if (SHOW_AS_BASE_FIELD.includes(show.as)) next.baseField = show.baseField ?? 0
  if (SHOW_AS_BASE_ITEM.includes(show.as)) next.baseItem = show.baseItem ?? "previous"
  return next
}

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
  let next: PivotDataField = { ...v, name, fn: settings.fn }
  if (settings.numFmt && settings.numFmt !== "General") next.numFmt = settings.numFmt
  else delete next.numFmt
  if (settings.show) next = withShowAs(next, settings.show)
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
  mergeLabels: boolean
  /** Characters per level, 0 to 127. */
  indent: number
  pageOverThenDown: boolean
  /** Report filter fields per column (per row, Over Then Down), 0 to 255; 0: all in one. */
  pageWrap: number
  preserveFormatting: boolean
  /** Totals & Filters. */
  rowGrandTotals: boolean
  colGrandTotals: boolean
  multipleFieldFilters: boolean
  customListSort: boolean
  /** Display. */
  showDrill: boolean
  showDataTips: boolean
  valuesRow: boolean
  gridDropZones: boolean
  showItems: boolean
  fieldListSortAscending: boolean
  /** Printing. */
  printDrill: boolean
  itemPrintTitles: boolean
  fieldPrintTitles: boolean
  /** Data (the last three belong to the cache: the caller's). */
  enableDrill: boolean
  refreshOnLoad: boolean
  saveData: boolean
  /** Number of items to retain per field: "auto", "none" or "max". */
  missingItems: "auto" | "none" | "max"
  /** Alt Text. */
  altText: string
  altTextSummary: string
}

/** Excel's Max for Number of items to retain per field (missingItemsLimit). */
export const PIVOT_MISSING_ITEMS_MAX = 1048576

/** The cache's missingItemsLimit as the dialog's choice, and back. */
export function pivotMissingItems(cache: PivotCache): PivotOptions["missingItems"] {
  if (cache.missingItemsLimit === undefined) return "auto"
  return cache.missingItemsLimit === 0 ? "none" : "max"
}
export const pivotMissingItemsLimit = (choice: PivotOptions["missingItems"]) =>
  choice === "auto" ? undefined : choice === "none" ? 0 : PIVOT_MISSING_ITEMS_MAX

/** The options a PivotTable has now, for the dialog. */
export function pivotOptions(table: PivotTable, cache: PivotCache): PivotOptions {
  return {
    name: table.name,
    showError: table.showError,
    errorCaption: table.captions.error ?? "",
    showMissing: table.showMissing,
    missingCaption: table.captions.missing ?? "",
    autoFit: table.autoFit !== false,
    mergeLabels: table.mergeLabels === true,
    indent: table.indent ?? 1,
    pageOverThenDown: table.pageOverThenDown === true,
    pageWrap: table.pageWrap ?? 0,
    preserveFormatting: table.preserveFormatting !== false,
    rowGrandTotals: table.rowGrandTotals,
    colGrandTotals: table.colGrandTotals,
    multipleFieldFilters: table.multipleFieldFilters === true,
    customListSort: table.customListSort !== false,
    showDrill: table.showDrill !== false,
    showDataTips: table.showDataTips !== false,
    valuesRow: !table.hideValuesRow,
    gridDropZones: table.gridDropZones === true,
    showItems: table.showItems !== false,
    fieldListSortAscending: table.fieldListSortAscending === true,
    printDrill: table.printDrill === true,
    itemPrintTitles: table.itemPrintTitles !== false,
    fieldPrintTitles: table.fieldPrintTitles === true,
    enableDrill: table.enableDrill !== false,
    refreshOnLoad: cache.refreshOnLoad,
    saveData: cache.saveData !== false,
    missingItems: pivotMissingItems(cache),
    altText: table.altText ?? "",
    altTextSummary: table.altTextSummary ?? "",
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
  // Options kept only when they differ from Excel's default: `false` ones default on, `true` ones off.
  type Off = "autoFit" | "showDrill" | "enableDrill" | "preserveFormatting" | "customListSort" | "showDataTips"
  type On = "mergeLabels" | "pageOverThenDown" | "multipleFieldFilters" | "fieldListSortAscending" | "printDrill"
  const flag = (key: Off | "showItems" | "itemPrintTitles", on: boolean) => {
    if (on) delete next[key]
    else next[key] = false
  }
  const onFlag = (key: On | "gridDropZones" | "fieldPrintTitles", on: boolean) => {
    if (on) next[key] = true
    else delete next[key]
  }
  for (const key of [
    "autoFit",
    "showDrill",
    "enableDrill",
    "preserveFormatting",
    "customListSort",
    "showDataTips",
    "showItems",
    "itemPrintTitles",
  ] as const)
    flag(key, options[key])
  for (const key of [
    "mergeLabels",
    "pageOverThenDown",
    "multipleFieldFilters",
    "fieldListSortAscending",
    "printDrill",
    "gridDropZones",
    "fieldPrintTitles",
  ] as const)
    onFlag(key, options[key])
  if (options.valuesRow) delete next.hideValuesRow
  else next.hideValuesRow = true
  const indent = Math.max(0, Math.min(127, Math.round(options.indent)))
  if (indent === 1) delete next.indent
  else next.indent = indent
  const wrap = Math.max(0, Math.min(255, Math.round(options.pageWrap)))
  if (wrap) next.pageWrap = wrap
  else delete next.pageWrap
  for (const key of ["altText", "altTextSummary"] as const) {
    if (options[key]) next[key] = options[key]
    else delete next[key]
  }
  // Classic PivotTable layout turned on puts every field in tabular form, as Excel's dialog does.
  return options.gridDropZones && !table.gridDropZones ? withReportLayout(next, "tabular") : next
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

/**
 * What Sort works on at a cell of a PivotTable (the cell menu's Sort, Data › Sort): a label sorts its field by the
 * labels (`by` null); a value cell sorts the field of its row by that value field (`by`, an index into `values`), or
 * the field of its column when it's in the Grand Total row. Null where Cells can't sort as Excel would: a value under
 * one column item (Excel sorts by that column alone; Cells sorts by the totals only).
 */
export function pivotSortTarget(
  table: PivotTable,
  cache: PivotCache,
  row: number,
  col: number,
): { field: number; by: number | null } | null {
  const layout = pivotLayout(table, cache)
  const cell = layout.cells.find((c) => c.row === row && c.col === col)
  if (!cell) return null
  if (cell.item) return { field: cell.item.field, by: null }
  if (cell.role === "rowHeader" || cell.role === "colHeader") {
    const field = pivotFilterTarget(table, cache, cell)?.fields.find((f) => f >= 0)
    return field === undefined ? null : { field, by: null }
  }
  if (!cell.lines) return null
  const target = pivotSettingsTarget(table, cache, row, col)
  if (target?.kind !== "value") return null
  const [ri, ci] = cell.lines
  const rowLine = layout.rowItems[ri]
  const colLine = layout.colItems[ci]
  const real = (axis: number[]) => axis.some((f) => f >= 0)
  /** The field a line belongs to: its deepest item's, stepping past the Values entry. */
  const fieldOf = (axis: number[], line: PivotAxisItem | undefined) => {
    if (!line || line.t === "grand" || line.t === "blank") return null
    for (let level = line.r + line.x.length - 1; level >= 0; level--) if (axis[level] >= 0) return axis[level]
    return null
  }
  const totalsOnly = (axis: number[], line: PivotAxisItem | undefined) => !real(axis) || !line || line.t === "grand"
  const rowField = fieldOf(table.rows, rowLine)
  if (rowField !== null && totalsOnly(table.cols, colLine)) return { field: rowField, by: target.index }
  const colField = fieldOf(table.cols, colLine)
  if (colField !== null && totalsOnly(table.rows, rowLine)) return { field: colField, by: target.index }
  return null
}

/**
 * Field `f` set to Manual sort keeping the order it shows now (a sort by value leaves the items in label order
 * underneath): its items stored in the order they first appear in the table, as Excel keeps them.
 */
export function withManualAsShown(table: PivotTable, cache: PivotCache, f: number): PivotTable {
  const field = table.fields[f]
  if (!field || field.sort === "manual") return table
  const shown: number[] = []
  for (const cell of pivotLayout(table, cache).cells)
    if (cell.item?.field === f && !shown.includes(cell.item.x)) shown.push(cell.item.x)
  const all = fullItems(field, fieldShared(cache, f).length)
  const rank = (x: number) => {
    const i = shown.indexOf(x)
    return i < 0 ? shown.length : i
  }
  const items = all.map((item, i) => ({ item, i })).sort((a, b) => rank(a.item.x) - rank(b.item.x) || a.i - b.i)
  const next: PivotField = { ...field, sort: "manual", items: items.map((e) => e.item) }
  delete next.sortBy
  const fields = [...table.fields]
  fields[f] = next
  return { ...table, fields }
}
