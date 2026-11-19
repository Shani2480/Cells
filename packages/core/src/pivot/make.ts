import type { CellAddress } from "../address"
import type { Sheet } from "../sheet"
import {
  newPivotField,
  PIVOT_FN_LABEL,
  VALUES_FIELD,
  type PivotCache,
  type PivotDataField,
  type PivotFn,
  type PivotTable,
} from "./model"

/**
 * Making and changing PivotTables (Insert › PivotTable, the PivotTable Fields pane): a new PivotTable as Excel 365
 * makes one, and the PivotTable with a field added to, moved within or taken out of the Filters, Columns, Rows and
 * Values areas. Laying the cells out again is the caller's.
 */

/** An area of the PivotTable Fields pane. */
export type PivotArea = "filters" | "cols" | "rows" | "values"

/** The next free PivotTableN in the workbook. */
export function nextPivotName(sheets: readonly Sheet[]): string {
  const taken = new Set(sheets.flatMap((s) => s.pivots.map((t) => t.name.toLowerCase())))
  let n = 1
  while (taken.has(`pivottable${n}`)) n++
  return `PivotTable${n}`
}

/** The next free cache id. */
export const nextPivotCacheId = (caches: readonly PivotCache[]) => Math.max(0, ...caches.map((c) => c.id)) + 1

/** An empty PivotTable at `at`, with Excel 365's defaults (compact form, PivotStyleLight16, no Values row). */
export function newPivotTable(name: string, cache: PivotCache, at: CellAddress): PivotTable {
  return {
    name,
    cacheId: cache.id,
    at,
    location: { r0: at.row, c0: at.col, r1: at.row, c1: at.col },
    fields: cache.fields.map((f) => newPivotField(f.items?.length ?? 0)),
    rows: [],
    cols: [],
    pages: [],
    values: [],
    rowGrandTotals: true,
    colGrandTotals: true,
    compact: true,
    outline: true,
    compactData: true,
    captions: { data: "Values" },
    showError: false,
    showMissing: true,
    hideValuesRow: true,
    style: {
      name: "PivotStyleLight16",
      rowHeaders: true,
      colHeaders: true,
      rowStripes: false,
      colStripes: false,
      lastColumn: true,
    },
  }
}

/** A PivotTable with no fields anywhere: Excel shows its placeholder instead of cells. */
export const pivotIsEmpty = (t: PivotTable) => !t.rows.length && !t.cols.length && !t.values.length

/** The areas a field is in (Values once per value field made from it). */
export function pivotFieldAreas(table: PivotTable, f: number): PivotArea[] {
  const out: PivotArea[] = []
  if (table.pages.some((p) => p.field === f)) out.push("filters")
  if (table.cols.includes(f)) out.push("cols")
  if (table.rows.includes(f)) out.push("rows")
  for (const v of table.values) if (v.field === f) out.push("values")
  return out
}

/** A field's records hold only numbers (and blanks): ticking it sums it. */
function numeric(cache: PivotCache, f: number): boolean {
  if (cache.fields[f]?.dates) return false
  let any = false
  for (const r of cache.records) {
    const v = r[f] ?? null
    if (v === null || v === "") continue
    if (typeof v !== "number") return false
    any = true
  }
  return any
}

/** A value field's own name, Excel's way: "Sum of Units", then "Sum of Units2" for a second one. */
function dataFieldName(table: PivotTable, cache: PivotCache, f: number, fn: PivotFn): string {
  const base = `${PIVOT_FN_LABEL[fn]} of ${table.fields[f]?.name ?? cache.fields[f]?.name ?? ""}`
  const taken = new Set(table.values.map((v) => v.name.toLowerCase()))
  if (!taken.has(base.toLowerCase())) return base
  let n = 2
  while (taken.has(`${base}${n}`.toLowerCase())) n++
  return `${base}${n}`
}

/** The Σ Values field goes into Columns once there are two value fields, and leaves when there is one. */
function withValuesField(table: PivotTable): PivotTable {
  const has = table.rows.includes(VALUES_FIELD) || table.cols.includes(VALUES_FIELD)
  if (table.values.length > 1 && !has) return { ...table, cols: [...table.cols, VALUES_FIELD] }
  if (table.values.length <= 1 && has)
    return {
      ...table,
      rows: table.rows.filter((f) => f !== VALUES_FIELD),
      cols: table.cols.filter((f) => f !== VALUES_FIELD),
    }
  return table
}

/** The fields' axis and data flags, from where they are. */
function withFlags(table: PivotTable): PivotTable {
  const fields = table.fields.map((field, f) => {
    const axis = table.rows.includes(f)
      ? ("row" as const)
      : table.cols.includes(f)
        ? ("col" as const)
        : table.pages.some((p) => p.field === f)
          ? ("page" as const)
          : null
    const dataField = table.values.some((v) => v.field === f)
    return field.axis === axis && field.dataField === dataField ? field : { ...field, axis, dataField }
  })
  return { ...table, fields }
}

/** The table with field `f` taken out of `area` (out of Values: value field `index`, or all made from it). */
export function withoutField(table: PivotTable, f: number, area: PivotArea, index?: number): PivotTable {
  let next: PivotTable
  if (area === "rows") next = { ...table, rows: table.rows.filter((x) => x !== f) }
  else if (area === "cols") next = { ...table, cols: table.cols.filter((x) => x !== f) }
  else if (area === "filters") next = { ...table, pages: table.pages.filter((p) => p.field !== f) }
  else if (f === VALUES_FIELD) return table
  else
    next = {
      ...table,
      values: table.values.filter((v, i) => (index !== undefined ? i !== index : v.field !== f)),
    }
  // A field leaving the row and column areas takes its filter with it, as in Excel.
  if (area !== "values" && next.filters?.some((x) => x.field === f)) {
    const filters = next.filters.filter((x) => x.field !== f)
    next = { ...next, filters }
    if (!filters.length) delete next.filters
  }
  return withFlags(withValuesField(next))
}

/** Unticking a field in the list: it leaves every area. */
export function withFieldRemoved(table: PivotTable, f: number): PivotTable {
  let next = table
  for (const area of ["filters", "cols", "rows", "values"] as const) next = withoutField(next, f, area)
  return next
}

/**
 * The table with field `f` put in `area` at `index` (the end when undefined). A field is on one of Filters,
 * Columns and Rows at a time, so it leaves the other two; Values takes a new value field each time (Sum for numbers,
 * Count otherwise). The Σ Values field (VALUES_FIELD) moves between Columns and Rows only.
 */
export function withFieldIn(
  table: PivotTable,
  cache: PivotCache,
  f: number,
  area: PivotArea,
  index?: number,
): PivotTable {
  const put = <T>(list: T[], item: T) => {
    const at = index === undefined ? list.length : Math.max(0, Math.min(index, list.length))
    return [...list.slice(0, at), item, ...list.slice(at)]
  }
  if (area === "values") {
    if (f === VALUES_FIELD) return table
    const fn: PivotFn = numeric(cache, f) ? "sum" : "count"
    const value: PivotDataField = { name: dataFieldName(table, cache, f, fn), field: f, fn }
    return withFlags(withValuesField({ ...table, values: put(table.values, value) }))
  }
  if (f === VALUES_FIELD && area === "filters") return table
  // Out of the axis areas first (a move within one keeps its place in the count).
  const rows = table.rows.filter((x) => x !== f)
  const cols = table.cols.filter((x) => x !== f)
  const pages = table.pages.filter((p) => p.field !== f)
  let next: PivotTable = { ...table, rows, cols, pages }
  if (area === "rows") next.rows = put(rows, f)
  else if (area === "cols") next.cols = put(cols, f)
  else next.pages = put(pages, { field: f, item: null })
  // Leaving rows and columns for Filters drops the field's label and value filters.
  if (area === "filters" && next.filters?.some((x) => x.field === f)) {
    const filters = next.filters.filter((x) => x.field !== f)
    next = { ...next, filters }
    if (!filters.length) delete next.filters
  }
  return withFlags(next)
}

/** Ticking a field in the list, as Excel does: numbers to Values, text and dates to Rows. */
export function withFieldTicked(table: PivotTable, cache: PivotCache, f: number): PivotTable {
  return withFieldIn(table, cache, f, numeric(cache, f) ? "values" : "rows")
}

/** Moving a value field to another place among the value fields. */
export function withValueMoved(table: PivotTable, from: number, to: number): PivotTable {
  const values = [...table.values]
  const [moved] = values.splice(from, 1)
  if (!moved) return table
  values.splice(Math.max(0, Math.min(to, values.length)), 0, moved)
  return { ...table, values }
}
