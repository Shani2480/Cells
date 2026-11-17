import {
  ICON_SETS,
  parseRange,
  type IconSetName,
  type AutoFilter,
  type ColumnFilter,
  type DateGroup,
  type DynamicFilter,
  type FilterCondition,
  type FilterOperator,
} from "@cells/core"
import { escapeXml } from "@cells/ooxml"
import { at, attrText, kid, kids, type XNode } from "./dom"
import type { StylesIn, StylesOut } from "./styles"

/**
 * A worksheet's `<autoFilter>`, as Excel writes it: the range, a `<filterColumn>` per filtered column (values with
 * date groups, custom conditions, Top 10, dynamic, colour) and the dropdown sort's `<sortState>`.
 */

const OPERATORS: Record<string, FilterOperator> = {
  equal: "eq",
  notEqual: "ne",
  greaterThan: "gt",
  greaterThanOrEqual: "ge",
  lessThan: "lt",
  lessThanOrEqual: "le",
}
const OPERATOR_NAMES: Record<FilterOperator, string | undefined> = {
  eq: undefined,
  ne: "notEqual",
  gt: "greaterThan",
  ge: "greaterThanOrEqual",
  lt: "lessThan",
  le: "lessThanOrEqual",
}

const DYNAMIC = new Set<string>([
  "aboveAverage",
  "belowAverage",
  "today",
  "yesterday",
  "tomorrow",
  "thisWeek",
  "lastWeek",
  "nextWeek",
  "thisMonth",
  "lastMonth",
  "nextMonth",
  "thisQuarter",
  "lastQuarter",
  "nextQuarter",
  "thisYear",
  "lastYear",
  "nextYear",
  "yearToDate",
  ...[1, 2, 3, 4].map((q) => `Q${q}`),
  ...Array.from({ length: 12 }, (_, i) => `M${i + 1}`),
])

const flag = (v: string | undefined, dflt: boolean) => (v === undefined ? dflt : v === "1" || v === "true")

function dateGroup(n: XNode): DateGroup | null {
  const year = Number(at(n, "year"))
  if (!Number.isInteger(year)) return null
  const grouping = at(n, "dateTimeGrouping") ?? "year"
  const g: DateGroup = { year }
  if (grouping !== "year") g.month = Number(at(n, "month"))
  if (grouping !== "year" && grouping !== "month") g.day = Number(at(n, "day"))
  return g
}

function readColumn(fc: XNode, styles: StylesIn | null, warnings: string[]): ColumnFilter | null {
  const filters = kid(fc, "filters")
  // Excel 365 writes date groups alone in an extension (xlrd2:filterColumn), with an empty filterColumn around it.
  const extFilters = kids(kid(fc, "extLst"), "ext")
    .map((e) => kid(kid(e, "filterColumn"), "filters"))
    .find(Boolean)
  if (filters || extFilters) {
    const values = kids(filters, "filter").map((f) => at(f, "val") ?? "")
    const dates = [...kids(filters, "dateGroupItem"), ...kids(extFilters, "dateGroupItem")]
      .map(dateGroup)
      .filter((g): g is DateGroup => g !== null)
    const out: ColumnFilter = { kind: "values", values }
    if (flag(at(filters ?? extFilters, "blank"), false)) out.blanks = true
    if (dates.length) out.dates = dates
    return out
  }
  const custom = kid(fc, "customFilters")
  if (custom) {
    const conditions: FilterCondition[] = kids(custom, "customFilter").map((c) => ({
      op: OPERATORS[at(c, "operator") ?? "equal"] ?? "eq",
      value: at(c, "val") ?? "",
    }))
    const out: ColumnFilter = { kind: "custom", conditions }
    if (conditions.length > 1 && !flag(at(custom, "and"), false)) out.and = false
    return out
  }
  const top = kid(fc, "top10")
  if (top)
    return {
      kind: "top10",
      top: flag(at(top, "top"), true),
      percent: flag(at(top, "percent"), false),
      count: Number(at(top, "val") ?? 10),
    }
  const dynamic = kid(fc, "dynamicFilter")
  if (dynamic) {
    const type = at(dynamic, "type") ?? ""
    if (DYNAMIC.has(type)) return { kind: "dynamic", type: type as DynamicFilter }
    warnings.push(`A date filter of type ${type} isn't supported; the filter on its column was left out.`)
    return null
  }
  const color = kid(fc, "colorFilter")
  if (color) {
    const cell = flag(at(color, "cellColor"), true)
    const dxf = at(color, "dxfId")
    return { kind: "color", cell, color: dxf !== undefined && styles ? styles.filterColor(Number(dxf), cell) : null }
  }
  const icon = kid(fc, "iconFilter")
  if (icon) {
    // iconId counts from the set's lowest band; no iconId is No Cell Icon.
    const set = at(icon, "iconSet") ?? "3Arrows"
    const id = at(icon, "iconId")
    if (set in ICON_SETS)
      return { kind: "icon", icon: id === undefined ? null : { set: set as IconSetName, index: Number(id) } }
    warnings.push(`A filter by the ${set} icons isn't supported; the filter on its column was left out.`)
  }
  return null
}

/** The sheet's AutoFilter from its `<autoFilter>` element (null when its range doesn't read). */
export function readAutoFilter(node: XNode, styles: StylesIn | null, warnings: string[]): AutoFilter | null {
  const range = parseRange(at(node, "ref") ?? "")
  if (!range) return null
  const filter: AutoFilter = { range, columns: {} }
  for (const fc of kids(node, "filterColumn")) {
    const col = Number(at(fc, "colId"))
    if (!Number.isInteger(col) || col < 0 || range.c0 + col > range.c1) continue
    const column = readColumn(fc, styles, warnings)
    if (column) filter.columns[col] = column
  }
  const condition = kid(kid(node, "sortState"), "sortCondition")
  const sortRef = condition ? parseRange(at(condition, "ref") ?? "") : null
  if (condition && sortRef && !at(condition, "sortBy"))
    filter.sort = { col: sortRef.c0 - range.c0, descending: flag(at(condition, "descending"), false) }
  return filter
}

const DATE_GROUPING = (g: DateGroup) => (g.day !== undefined ? "day" : g.month !== undefined ? "month" : "year")

function columnXml(column: ColumnFilter, styles: StylesOut): string {
  switch (column.kind) {
    case "values": {
      const items = column.values.map((v) => `<filter val="${escapeXml(v)}"/>`)
      for (const g of column.dates ?? [])
        items.push(
          `<dateGroupItem${attrText({ year: g.year, month: g.month, day: g.day, dateTimeGrouping: DATE_GROUPING(g) })}/>`,
        )
      const open = `<filters${column.blanks ? ` blank="1"` : ""}`
      return items.length ? `${open}>${items.join("")}</filters>` : `${open}/>`
    }
    case "custom":
      return (
        `<customFilters${column.and !== false && column.conditions.length > 1 ? ` and="1"` : ""}>` +
        column.conditions
          .map((c) => `<customFilter${attrText({ operator: OPERATOR_NAMES[c.op], val: c.value })}/>`)
          .join("") +
        `</customFilters>`
      )
    case "top10":
      return `<top10${attrText({ top: column.top ? undefined : "0", percent: column.percent || undefined, val: column.count })}/>`
    case "dynamic":
      return `<dynamicFilter type="${column.type}"/>`
    case "color":
      return `<colorFilter${attrText({ dxfId: styles.filterDxf(column.cell, column.color), cellColor: column.cell ? undefined : "0" })}/>`
    case "icon":
      // No Cell Icon names no set: Excel writes the column's set; 3Arrows reads back as the same filter.
      return `<iconFilter${attrText({ iconSet: column.icon?.set ?? "3Arrows", iconId: column.icon?.index })}/>`
  }
}

/** `<autoFilter>` for a sheet's filter. */
export function autoFilterXml(filter: AutoFilter, styles: StylesOut, rangeRef: (r: AutoFilter["range"]) => string) {
  const columns = Object.entries(filter.columns)
    .sort((a, b) => Number(a[0]) - Number(b[0]))
    .map(([col, column]) => `<filterColumn colId="${col}">${columnXml(column, styles)}</filterColumn>`)
  const { range, sort } = filter
  if (sort && range.r1 > range.r0) {
    const data = { ...range, r0: range.r0 + 1 }
    const c = range.c0 + sort.col
    columns.push(
      `<sortState ref="${rangeRef(data)}"><sortCondition${sort.descending ? ` descending="1"` : ""} ref="${rangeRef({ ...data, c0: c, c1: c })}"/></sortState>`,
    )
  }
  const ref = rangeRef(range)
  return columns.length ? `<autoFilter ref="${ref}">${columns.join("")}</autoFilter>` : `<autoFilter ref="${ref}"/>`
}
