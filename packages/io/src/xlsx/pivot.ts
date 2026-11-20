import {
  errorValue,
  discreteLevels,
  fullItems,
  inRecords,
  parseRange,
  pivotLayout,
  rangeName,
  SHOW_AS_BASE_FIELD,
  SHOW_AS_BASE_ITEM,
  VALUES_FIELD,
  type ErrorCode,
  type PivotAxisItem,
  type PivotCache,
  type PivotCacheField,
  type PivotField,
  type PivotFilter,
  type PivotLabelOp,
  type PivotValueOp,
  type PivotFieldGroup,
  type PivotFn,
  type PivotGroupBy,
  type PivotShowAs,
  type PivotSort,
  type PivotSubtotal,
  type PivotTable,
  type PivotValue,
} from "@cells/core"
import { escapeXml, parseXmlEvents } from "@cells/ooxml"
import { NS_MAIN, NS_R } from "./constants"
import { at, attrText, kid, kids, parseTree, raw, type Kept, type XNode } from "./dom"

/**
 * PivotTables in xlsx: the cache definition (fields and shared items), the cache records, and each PivotTable's
 * definition (ECMA-376 Part 1, §18.10). What Cells models is read into wb.pivotCaches and sheet.pivots; a cache that
 * uses what Cells doesn't model yet (OLAP, external data, grouping, calculated fields or items, Show Values As,
 * date filters) is kept as read with all its PivotTables, as before.
 *
 * On save a cache or PivotTable that hasn't changed is written back byte for byte; one that only moved (rows
 * inserted above it) gets its location patched; a changed one is written anew.
 */

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
export const REL_PIVOT_CACHE = REL + "pivotCacheDefinition"
export const REL_PIVOT_RECORDS = REL + "pivotCacheRecords"
export const REL_PIVOT_TABLE = REL + "pivotTable"
const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml."
export const CT_PIVOT_CACHE = CT + "pivotCacheDefinition+xml"
export const CT_PIVOT_RECORDS = CT + "pivotCacheRecords+xml"
export const CT_PIVOT_TABLE = CT + "pivotTable+xml"

const ISO_EPOCH = Date.UTC(1899, 11, 30)
const DAY = 86_400_000

const truthy = (v: string | undefined, fallback: boolean) => (v === undefined ? fallback : v === "1" || v === "true")

export function isoToSerial(text: string): number | null {
  const ms = Date.parse(/Z$|[+-]\d\d:\d\d$/.test(text) ? text : text + "Z")
  return Number.isNaN(ms) ? null : (ms - ISO_EPOCH) / DAY
}

export function serialToIso(serial: number): string {
  return new Date(Math.round(ISO_EPOCH + serial * DAY)).toISOString().replace(/\.\d+Z$/, "")
}

/** One cache value element (s, n, d, b, e, m). */
function valueOf(local: string, v: string | undefined): PivotValue | undefined {
  switch (local) {
    case "s":
      return v ?? ""
    case "n":
      return Number(v)
    case "d":
      return v !== undefined ? isoToSerial(v) : null
    case "b":
      return v === "1" || v === "true"
    case "e":
      return errorValue((v ?? "#N/A") as ErrorCode)
    case "m":
      return null
    default:
      return undefined
  }
}

// ── Reading ──

/** A cache definition and its records; null when it uses what Cells doesn't model. */
export function readPivotCache(
  id: number,
  defXml: string,
  recordsXml: string | undefined,
  numberFormat: (id: number) => string | undefined,
): PivotCache | null {
  const root = parseTree(defXml).root
  const src = kid(root, "cacheSource")
  if (!src || (at(src, "type") ?? "worksheet") !== "worksheet") return null
  const ws = kid(src, "worksheetSource")
  if (!ws || at(ws, "id", NS_R)) return null
  if (kid(root, "cacheHierarchies") || kid(root, "calculatedItems") || kid(root, "calculatedMembers")) return null
  let source: PivotCache["source"]
  const name = at(ws, "name")
  if (name) source = { kind: "name", name }
  else {
    const ref = at(ws, "ref")
    const sheet = at(ws, "sheet")
    if (!ref || !sheet || !parseRange(ref)) return null
    source = { kind: "range", sheet, ref }
  }
  const fields: PivotCacheField[] = []
  const cacheFields = kids(kid(root, "cacheFields"), "cacheField")
  for (const f of cacheFields) {
    const group = readFieldGroup(kid(f, "fieldGroup"), fields.length)
    if (group === null) return null
    // A calculated field: a formula and no data; a group field of its own: no data either. Excel puts them after the
    // source's columns, as the records need.
    const formula = at(f, "formula")
    if (formula !== undefined || at(f, "databaseField") === "0") {
      if (formula === undefined && (!group || group.base === fields.length)) return null
      const calc: PivotCacheField =
        formula !== undefined
          ? { name: at(f, "name") ?? "", items: null, formula }
          : { name: at(f, "name") ?? "", items: null, group: group! }
      const fmt = numberFormat(Number(at(f, "numFmtId") ?? 0))
      if (fmt && fmt !== "General") calc.numFmt = fmt
      fields.push(calc)
      continue
    }
    if (fields.some((x, i) => !inRecords(x, i))) return null
    if (group && group.base !== fields.length) return null
    const shared = kid(f, "sharedItems")
    const listed = shared?.children.filter((c) => "snbdem".includes(c.local) && c.local.length === 1) ?? []
    const items = listed.length ? listed.map((c) => valueOf(c.local, at(c, "v")) ?? null) : null
    const field: PivotCacheField = { name: at(f, "name") ?? "", items }
    const fmt = numberFormat(Number(at(f, "numFmtId") ?? 0))
    if (fmt && fmt !== "General") field.numFmt = fmt
    if (truthy(at(shared, "containsDate"), false)) field.dates = true
    if (group) field.group = group
    fields.push(field)
  }
  const records = recordsXml ? readRecords(recordsXml, fields) : []
  const cache: PivotCache = { id, source, fields, records, refreshOnLoad: truthy(at(root, "refreshOnLoad"), false) }
  const x14 = kids(kid(root, "extLst"), "ext").find((e) => at(e, "uri") === X14_CACHE)
  const slicerId = Number(at(kid(x14, "pivotCacheDefinition"), "pivotCacheId"))
  if (x14 && Number.isFinite(slicerId)) cache.slicerId = slicerId
  return cache
}

const GROUP_BY = new Set(["range", "seconds", "minutes", "hours", "days", "months", "quarters", "years"])

/**
 * A cache field's grouping: undefined when it has none of its own (no fieldGroup, or only `par`, which points at the
 * field grouped from it); null when it is one Cells can't read.
 */
function readFieldGroup(fg: XNode | undefined, index: number): PivotFieldGroup | undefined | null {
  if (!fg) return undefined
  const range = kid(fg, "rangePr")
  const discrete = kid(fg, "discretePr")
  if (!range && !discrete) return at(fg, "base") === undefined ? undefined : null
  const base = Number(at(fg, "base") ?? index)
  const items = (kid(fg, "groupItems")?.children ?? []).map((c) => valueOf(c.local, at(c, "v")) ?? null)
  if (discrete) return { kind: "discrete", base, map: kids(discrete, "x").map((x) => Number(at(x, "v") ?? 0)), items }
  const by = at(range, "groupBy") ?? "range"
  if (!GROUP_BY.has(by)) return null
  const dates = by !== "range"
  const bound = (name: string, fallback: number) => {
    const v = dates ? valueOf("d", at(range, `${name}Date`)) : Number(at(range, `${name}Num`) ?? NaN)
    return typeof v === "number" && Number.isFinite(v) ? v : fallback
  }
  return {
    kind: "range",
    base,
    by: by as PivotGroupBy,
    start: bound("start", 0),
    end: bound("end", 0),
    interval: Number(at(range, "groupInterval") ?? 1),
    autoStart: truthy(at(range, "autoStart"), true),
    autoEnd: truthy(at(range, "autoEnd"), true),
    items,
  }
}

/** The records, streamed (a cache can hold many thousands of rows). */
function readRecords(xml: string, fields: PivotCacheField[]): PivotValue[][] {
  const out: PivotValue[][] = []
  let row: PivotValue[] | null = null
  parseXmlEvents(xml, {
    start(tag) {
      if (tag.local === "r") row = []
      else if (row) {
        const v = tag.attributes.find((a) => a.local === "v")?.value
        if (tag.local === "x") {
          const items = fields[row.length]?.items
          row.push(items?.[Number(v ?? 0)] ?? null)
        } else {
          const value = valueOf(tag.local, v)
          if (value !== undefined) row.push(value)
        }
      }
    },
    end(tag) {
      if (tag.local === "r" && row) {
        out.push(row)
        row = null
      }
    },
    text() {},
  })
  return out
}

const ITEM_FN: Record<string, PivotFn> = {
  sum: "sum",
  countA: "count",
  avg: "average",
  max: "max",
  min: "min",
  product: "product",
  count: "countNums",
  stdDev: "stdDev",
  stdDevP: "stdDevp",
  var: "var",
  varP: "varp",
}
const FN_ITEM = Object.fromEntries(Object.entries(ITEM_FN).map(([k, v]) => [v, k])) as Record<PivotFn, string>
/** Field attributes that turn on each custom subtotal, in the order Excel's Field Settings lists them. */
const FN_ATTR: [PivotFn, string][] = [
  ["sum", "sumSubtotal"],
  ["count", "countASubtotal"],
  ["average", "avgSubtotal"],
  ["max", "maxSubtotal"],
  ["min", "minSubtotal"],
  ["product", "productSubtotal"],
  ["countNums", "countSubtotal"],
  ["stdDev", "stdDevSubtotal"],
  ["stdDevp", "stdDevPSubtotal"],
  ["var", "varSubtotal"],
  ["varp", "varPSubtotal"],
]
/** dataField's subtotal attribute. */
const DATA_FN: Record<string, PivotFn> = {
  sum: "sum",
  count: "count",
  average: "average",
  max: "max",
  min: "min",
  product: "product",
  countNums: "countNums",
  stdDev: "stdDev",
  stdDevp: "stdDevp",
  var: "var",
  varp: "varp",
}
const AXIS: Record<string, PivotField["axis"]> = { axisRow: "row", axisCol: "col", axisPage: "page" }
const X14_FIELD = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}"
const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

/** The x14 extension of a dataField (Show Values As choices added in Excel 2010). */
const X14_DATA = "{E15A36E0-9728-4e99-A89B-3F7291B0FE68}"
/** baseItem's (previous) and (next). */
const BASE_PREVIOUS = 1048828
const BASE_NEXT = 1048829
/** The showDataAs values of the file format; the rest go in the x14 extension. */
const SHOW_DATA_AS = new Set([
  "difference",
  "percent",
  "percentDiff",
  "runTotal",
  "percentOfRow",
  "percentOfCol",
  "percentOfTotal",
  "index",
])
const SHOW_AS = new Set([
  ...SHOW_DATA_AS,
  "percentOfParentRow",
  "percentOfParentCol",
  "percentOfParent",
  "percentOfRunningTotal",
  "rankAscending",
  "rankDescending",
])

/** A PivotTable definition; null when it uses what Cells doesn't model. */
export function readPivotTable(xml: string, numberFormat: (id: number) => string | undefined): PivotTable | null {
  const root = parseTree(xml).root
  if (kid(root, "pivotHierarchies") || kid(root, "calculatedItems")) return null
  const filters = readPivotFilters(kid(root, "filters"))
  if (!filters) return null
  const loc = kid(root, "location")
  const location = parseRange(at(loc, "ref") ?? "")
  if (!location) return null
  const fmt = (n: XNode | undefined) => {
    const id = at(n, "numFmtId")
    const code = id !== undefined ? numberFormat(Number(id)) : undefined
    return code && code !== "General" ? code : undefined
  }
  const compact = truthy(at(root, "compact"), true)
  const outline = truthy(at(root, "outline"), false)
  const fields: PivotField[] = []
  for (const f of kids(kid(root, "pivotFields"), "pivotField")) {
    if (at(f, "autoShow") === "1" || kid(f, "autoSortScope") || at(f, "dataSourceSort") === "1") return null
    const axisAttr = at(f, "axis")
    if (axisAttr && !(axisAttr in AXIS)) return null
    const items: PivotField["items"] = []
    const fromItems: PivotSubtotal[] = []
    for (const item of kids(kid(f, "items"), "item")) {
      const t = at(item, "t")
      if (t === undefined || t === "data") {
        if (t === "data" || at(item, "f") === "1") return null
        const it: PivotField["items"][number] = { x: Number(at(item, "x") ?? 0) }
        if (truthy(at(item, "h"), false)) it.hidden = true
        if (!truthy(at(item, "sd"), true)) it.collapsed = true
        const n = at(item, "n")
        if (n !== undefined) it.name = n
        items.push(it)
      } else if (t === "default") fromItems.push("default")
      else if (t in ITEM_FN) fromItems.push(ITEM_FN[t])
      else if (t !== "blank" && t !== "grand") return null
    }
    const custom = FN_ATTR.filter(([, a]) => truthy(at(f, a), false)).map(([fn]) => fn)
    const subtotals: PivotSubtotal[] = custom.length
      ? fromItems.filter((s) => s !== "default").length === custom.length
        ? fromItems.filter((s) => s !== "default")
        : custom
      : truthy(at(f, "defaultSubtotal"), true)
        ? ["default"]
        : []
    const x14 = kids(kid(f, "extLst"), "ext").find((e) => at(e, "uri") === X14_FIELD)
    const field: PivotField = {
      axis: axisAttr ? AXIS[axisAttr] : null,
      dataField: truthy(at(f, "dataField"), false),
      items,
      subtotals,
      subtotalTop: truthy(at(f, "subtotalTop"), true),
      outline: truthy(at(f, "outline"), true),
      compact: truthy(at(f, "compact"), true),
      insertBlankRow: truthy(at(f, "insertBlankRow"), false),
      showAll: truthy(at(f, "showAll"), true),
      sort: (at(f, "sortType") as PivotSort | undefined) ?? "manual",
      repeatLabels: truthy(at(kid(x14, "pivotField"), "fillDownLabels"), false),
      multipleItems: truthy(at(f, "multipleItemSelectionAllowed"), false),
    }
    if (!["manual", "ascending", "descending"].includes(field.sort)) return null
    const name = at(f, "name")
    if (name !== undefined) field.name = name
    const numFmt = fmt(f)
    if (numFmt) field.numFmt = numFmt
    fields.push(field)
  }
  const list = (local: string) => kids(kid(root, local), "field").map((n) => Number(at(n, "x") ?? 0))
  const values: PivotTable["values"] = []
  const baseItems: [PivotTable["values"][number], number][] = []
  for (const d of kids(kid(root, "dataFields"), "dataField")) {
    // Show Values As: showDataAs, or for Excel 2010's choices the x14 extension (any other extension isn't known).
    const exts = kids(kid(d, "extLst"), "ext")
    const x14 = exts.find((e) => at(e, "uri") === X14_DATA)
    if (exts.length > (x14 ? 1 : 0)) return null
    const x14As = at(kid(x14, "dataField"), "pivotShowAs")
    const show = x14As ?? at(d, "showDataAs") ?? "normal"
    if (show !== "normal" && !SHOW_AS.has(show)) return null
    const fn = DATA_FN[at(d, "subtotal") ?? "sum"]
    if (!fn) return null
    const v: PivotTable["values"][number] = { name: at(d, "name") ?? "", field: Number(at(d, "fld") ?? 0), fn }
    const numFmt = fmt(d)
    if (numFmt) v.numFmt = numFmt
    if (show !== "normal") {
      v.showAs = show as PivotShowAs
      if (SHOW_AS_BASE_FIELD.includes(v.showAs)) v.baseField = Number(at(d, "baseField") ?? 0)
      if (SHOW_AS_BASE_ITEM.includes(v.showAs)) baseItems.push([v, Number(at(d, "baseItem") ?? 0)])
    }
    values.push(v)
  }
  // A base item is a position in the base field's items (the cache item it shows), or Excel's previous / next.
  for (const [v, item] of baseItems) {
    if (item === BASE_PREVIOUS) v.baseItem = "previous"
    else if (item === BASE_NEXT) v.baseItem = "next"
    else {
      const x = fields[v.baseField ?? 0]?.items[item]?.x
      if (x === undefined) return null
      v.baseItem = x
    }
  }
  const pages = kids(kid(root, "pageFields"), "pageField").map((p) => {
    const item = at(p, "item")
    return { field: Number(at(p, "fld") ?? 0), item: item !== undefined ? Number(item) : null }
  })
  const style = kid(root, "pivotTableStyleInfo")
  const caption = (name: string) => at(root, name)
  const captions: PivotTable["captions"] = { data: caption("dataCaption") ?? "Values" }
  for (const [key, attr] of [
    ["grandTotal", "grandTotalCaption"],
    ["rowHeader", "rowHeaderCaption"],
    ["colHeader", "colHeaderCaption"],
    ["error", "errorCaption"],
    ["missing", "missingCaption"],
  ] as const) {
    const v = caption(attr)
    if (v !== undefined) captions[key] = v
  }
  return {
    name: at(root, "name") ?? "PivotTable1",
    cacheId: Number(at(root, "cacheId") ?? 0),
    at: { row: location.r0, col: location.c0 },
    location,
    fields,
    rows: list("rowFields"),
    cols: list("colFields"),
    pages,
    values,
    rowGrandTotals: truthy(at(root, "rowGrandTotals"), true),
    colGrandTotals: truthy(at(root, "colGrandTotals"), true),
    compact,
    outline,
    compactData: truthy(at(root, "compactData"), true),
    captions,
    showError: truthy(at(root, "showError"), false),
    showMissing: truthy(at(root, "showMissing"), true),
    ...(truthy(at(root, "showDrill"), true) ? {} : { showDrill: false as const }),
    ...(truthy(at(root, "enableDrill"), true) ? {} : { enableDrill: false as const }),
    ...(truthy(at(root, "useAutoFormatting"), false) ? {} : { autoFit: false as const }),
    ...(filters.length ? { filters } : {}),
    ...(hideValuesRow(root) ? { hideValuesRow: true as const } : {}),
    style: {
      name: at(style, "name") ?? null,
      rowHeaders: truthy(at(style, "showRowHeaders"), false),
      colHeaders: truthy(at(style, "showColHeaders"), false),
      rowStripes: truthy(at(style, "showRowStripes"), false),
      colStripes: truthy(at(style, "showColStripes"), false),
      lastColumn: truthy(at(style, "showLastColumn"), false),
    },
  }
}

const X14_TABLE = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"

/** Excel 365's Show the Values row off, kept in the x14 extension of the definition (written back with the extLst). */
function hideValuesRow(root: XNode): boolean {
  const ext = kids(kid(root, "extLst"), "ext").find((e) => at(e, "uri") === X14_TABLE)
  return truthy(at(kid(ext, "pivotTableDefinition"), "hideValuesRow"), false)
}

// ── Label Filters, Value Filters and Top 10 (`<filters>`) ──

const LABEL_OPS: PivotLabelOp[] = [
  "equal",
  "notEqual",
  "beginsWith",
  "notBeginsWith",
  "endsWith",
  "notEndsWith",
  "contains",
  "notContains",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
  "between",
  "notBetween",
]
const VALUE_OPS: PivotValueOp[] = [
  "equal",
  "notEqual",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
  "between",
  "notBetween",
]
const cap = (op: string) => op[0].toUpperCase() + op.slice(1)

/** A date filter on whole days (Excel 2013's x15 pivotFilter useWholeDay, which timelines write). */
const wholeDay = (f: XNode) =>
  kids(kid(f, "extLst"), "ext").some((e) => truthy(at(kid(e, "pivotFilter"), "useWholeDay"), false))

/** The filters of a PivotTable; null when one is of a kind Cells doesn't model (date filters, OLAP members). */
function readPivotFilters(node: XNode | undefined): PivotFilter[] | null {
  const out: PivotFilter[] = []
  for (const f of kids(node, "filter")) {
    const type = at(f, "type") ?? ""
    const field = Number(at(f, "fld") ?? -1)
    const id = Number(at(f, "id") ?? out.length + 1)
    if (field < 0 || at(f, "mpFld") !== undefined || at(f, "iMeasureHier") !== undefined) return null
    const label = LABEL_OPS.find((op) => type === `caption${cap(op)}`)
    const value = VALUE_OPS.find((op) => type === `value${cap(op)}`)
    const measure = Number(at(f, "iMeasureFld") ?? 0)
    if (label) {
      const filter: PivotFilter = { field, id, kind: "label", op: label, value: at(f, "stringValue1") ?? "" }
      const v2 = at(f, "stringValue2")
      if (v2 !== undefined) filter.value2 = v2
      out.push(filter)
    } else if (value) {
      // The numbers are in the autoFilter's custom filters.
      const vals = kids(kid(kid(kid(f, "autoFilter"), "filterColumn"), "customFilters"), "customFilter").map((c) =>
        Number(at(c, "val")),
      )
      if (!vals.length || vals.some((v) => !Number.isFinite(v))) return null
      const filter: PivotFilter = { field, id, kind: "value", op: value, measure, value: vals[0] }
      if (vals.length > 1) filter.value2 = vals[1]
      out.push(filter)
    } else if (type === "dateBetween") {
      // A timeline's: whole days (Excel 2013's useWholeDay), from and to in the custom filters.
      const vals = kids(kid(kid(kid(f, "autoFilter"), "filterColumn"), "customFilters"), "customFilter").map((c) =>
        Number(at(c, "val")),
      )
      if (vals.length !== 2 || vals.some((v) => !Number.isFinite(v)) || !wholeDay(f)) return null
      out.push({ field, id, kind: "date", op: "between", value: vals[0], value2: vals[1] })
    } else if (type === "count" || type === "percent" || type === "sum") {
      const top10 = kid(kid(kid(f, "autoFilter"), "filterColumn"), "top10")
      const v = Number(at(top10, "val"))
      if (!top10 || !Number.isFinite(v)) return null
      out.push({ field, id, kind: "top", by: type, top: truthy(at(top10, "top"), true), measure, value: v })
    } else return null
  }
  return out
}

/** Excel's `<filter>` for one of a PivotTable's filters, with the autoFilter Excel writes beside it. */
function pivotFilterXml(f: PivotFilter): string {
  const custom = (list: { operator?: string; val: string | number }[], and = false) =>
    `<customFilters${and ? ' and="1"' : ""}>${list
      .map((c) => `<customFilter${attrText({ operator: c.operator, val: c.val })}/>`)
      .join("")}</customFilters>`
  let type: string
  let inner: string
  const extra: Record<string, string | number | undefined> = {}
  const range = (a: string | number, b: string | number, not: boolean) =>
    not
      ? custom([
          { operator: "lessThan", val: a },
          { operator: "greaterThan", val: b },
        ])
      : custom(
          [
            { operator: "greaterThanOrEqual", val: a },
            { operator: "lessThanOrEqual", val: b },
          ],
          true,
        )
  if (f.kind === "label") {
    type = `caption${cap(f.op)}`
    extra.stringValue1 = f.value
    extra.stringValue2 = f.op === "between" || f.op === "notBetween" ? (f.value2 ?? f.value) : undefined
    const lit = f.value.replace(/[~*?]/g, (c) => `~${c}`)
    const like: Partial<Record<PivotLabelOp, string>> = {
      equal: f.value,
      notEqual: f.value,
      beginsWith: `${lit}*`,
      notBeginsWith: `${lit}*`,
      endsWith: `*${lit}`,
      notEndsWith: `*${lit}`,
      contains: `*${lit}*`,
      notContains: `*${lit}*`,
    }
    const pattern = like[f.op]
    if (pattern !== undefined)
      inner = custom([{ operator: f.op.startsWith("not") ? "notEqual" : undefined, val: pattern }])
    else if (f.op === "between" || f.op === "notBetween")
      inner = range(f.value, f.value2 ?? f.value, f.op === "notBetween")
    else inner = custom([{ operator: f.op, val: f.value }])
  } else if (f.kind === "value") {
    type = `value${cap(f.op)}`
    extra.iMeasureFld = f.measure
    if (f.op === "between" || f.op === "notBetween") inner = range(f.value, f.value2 ?? f.value, f.op === "notBetween")
    else inner = custom([{ operator: f.op === "equal" ? undefined : f.op, val: f.value }])
  } else if (f.kind === "date") {
    const inner = range(f.value, f.value2, false)
    const head = attrText({ fld: f.field, type: "dateBetween", evalOrder: -1, id: f.id })
    return (
      `<filter${head}><autoFilter ref="A1"><filterColumn colId="0">${inner}</filterColumn></autoFilter>` +
      `<extLst><ext uri="${X15_FILTER}" xmlns:x15="${NS_X15}"><x15:pivotFilter useWholeDay="1"/></ext></extLst></filter>`
    )
  } else {
    type = f.by
    extra.iMeasureFld = f.measure
    inner = `<top10${attrText({
      top: f.top ? undefined : 0,
      percent: f.by === "percent" ? 1 : undefined,
      val: f.value,
      filterVal: f.value,
    })}/>`
  }
  const head = attrText({ fld: f.field, type, evalOrder: -1, id: f.id, iMeasureFld: extra.iMeasureFld })
  const strings = attrText({ stringValue1: extra.stringValue1, stringValue2: extra.stringValue2 })
  return `<filter${head}${strings}><autoFilter ref="A1"><filterColumn colId="0">${inner}</filterColumn></autoFilter></filter>`
}

/** The children of a PivotTable definition that Cells doesn't write itself, kept when it writes the table anew. */
export function keptPivotChildren(xml: string): Kept[] {
  const doc = parseTree(xml)
  return doc.root.children
    .filter((c) => ["formats", "conditionalFormats", "chartFormats", "extLst"].includes(c.local))
    .map((c) => ({ local: c.local, xml: raw(doc, c) }))
}

/** The table moved (rows or columns inserted or deleted before it): only its location changes. */
export function movedPivotXml(xml: string, table: PivotTable): string {
  return xml
    .replace(/^\uFEFF?<\?xml[^>]*\?>\s*/, "")
    .replace(/(<(?:\w+:)?location\b[^>]*?\bref=")[^"]*"/, `$1${rangeName(table.location)}"`)
}

/** What decides whether a PivotTable changed since it was read (its place aside). */
export const pivotTableKey = (t: PivotTable) => JSON.stringify({ ...t, at: null, location: null })
export const pivotPlaceKey = (t: PivotTable) => JSON.stringify([t.at, t.location])
/** What decides whether a cache changed since it was read (its records are compared by identity). */
export const pivotCacheKey = (c: PivotCache) => JSON.stringify({ ...c, records: null })

// ── Writing ──

const valueKey = (v: PivotValue): string =>
  v === null || v === ""
    ? "m"
    : typeof v === "number"
      ? `n:${v}`
      : typeof v === "string"
        ? `s:${v.toLowerCase()}`
        : typeof v === "boolean"
          ? `b:${v}`
          : `e:${v.error}`

/**
 * The cache as it is written: every field that a PivotTable puts on an axis lists its shared items (a number column
 * read without them gets its distinct values, in the order first met, as the engine counts them).
 */
export function cacheForWriting(cache: PivotCache, tables: readonly PivotTable[]): PivotCache {
  const needed = new Set<number>()
  for (const t of tables)
    t.fields.forEach((f, i) => {
      if (f.axis || f.items.length) needed.add(i)
    })
  // A grouped field's base lists its shared items, as Excel writes it.
  cache.fields.forEach((f) => f.group && needed.add(f.group.base))
  const fields = cache.fields.map((f, i) => {
    if ((!needed.has(i) && !f.items) || !inRecords(f, i)) return f
    const items = f.items ? [...f.items] : []
    const seen = new Set(items.map(valueKey))
    for (const rec of cache.records) {
      const v = rec[i] ?? null
      const k = valueKey(v)
      if (!seen.has(k)) {
        seen.add(k)
        items.push(v)
      }
    }
    return f.items && items.length === f.items.length ? f : { ...f, items }
  })
  return { ...cache, fields }
}

/**
 * A cache field's fieldGroup: `par` the last field grouped from it (Excel's choice), and its own grouping — number or
 * date ranges (rangePr) or items gathered by Group Selection (discretePr) — with the group items.
 */
function fieldGroupXml(cache: PivotCache, f: number): string {
  let par: number | undefined
  // A date's part fields: the last one. Group Selection's levels: the source points at the first, each at the next.
  cache.fields.forEach((k, i) => {
    if (i !== f && k.group?.kind === "range" && k.group.base === f) par = i
  })
  const own = cache.fields[f].group
  const root = own?.kind === "discrete" ? own.base : f
  const levels = discreteLevels(cache, root)
  const next = root === f ? levels[0] : levels[levels.indexOf(f) + 1]
  if (next !== undefined) par = next
  const g = cache.fields[f].group
  if (!g) return par === undefined ? "" : `<fieldGroup par="${par}"/>`
  const items = g.items.map((v) => valueXml(v, false)).join("")
  const groupItems = `<groupItems count="${g.items.length}">${items}</groupItems>`
  let inner: string
  if (g.kind === "discrete")
    inner = `<discretePr count="${g.map.length}">${g.map.map((x) => `<x v="${Math.max(x, 0)}"/>`).join("")}</discretePr>`
  else {
    const dates = g.by !== "range"
    inner =
      `<rangePr` +
      attrText({
        autoStart: g.autoStart ? undefined : 0,
        autoEnd: g.autoEnd ? undefined : 0,
        groupBy: dates ? g.by : undefined,
        startNum: dates ? undefined : g.start,
        endNum: dates ? undefined : g.end,
        startDate: dates ? serialToIso(g.start) : undefined,
        endDate: dates ? serialToIso(g.end) : undefined,
        groupInterval: !dates || g.interval !== 1 ? g.interval : undefined,
      }) +
      `/>`
  }
  return `<fieldGroup${attrText({ par, base: g.base })}>${inner}${groupItems}</fieldGroup>`
}

/** How many items a field's indexes can point at: its group items, or the cache's shared items it keeps. */
const sharedCount = (cache: PivotCache, f: number) => {
  const field = cache.fields[f]
  return field?.group ? field.group.items.length : (field?.items?.length ?? 0)
}

/** The value element for a cache value. */
function valueXml(v: PivotValue, dates: boolean): string {
  if (v === null || v === "") return v === "" ? `<s v=""/>` : "<m/>"
  if (typeof v === "number") return dates ? `<d v="${serialToIso(v)}"/>` : `<n v="${v}"/>`
  if (typeof v === "string") return `<s v="${escapeXml(v)}"/>`
  if (typeof v === "boolean") return `<b v="${v ? 1 : 0}"/>`
  return `<e v="${escapeXml(v.error)}"/>`
}

/** sharedItems' attributes, worked out from the values the way Excel writes them. */
function sharedItemsAttrs(values: Iterable<PivotValue>, dates: boolean) {
  let text = false
  let blank = false
  let number = false
  let date = false
  let integer = true
  let long = false
  let min = Infinity
  let max = -Infinity
  const kinds = new Set<string>()
  for (const v of values) {
    if (v === null || v === "") {
      blank = true
      continue
    }
    if (typeof v === "number") {
      if (dates) date = true
      else number = true
      kinds.add(dates ? "d" : "n")
      if (!Number.isInteger(v)) integer = false
      if (v < min) min = v
      if (v > max) max = v
    } else {
      text = true
      kinds.add(typeof v === "object" ? "e" : typeof v)
      if (typeof v === "string" && v.length > 255) long = true
    }
  }
  return attrText({
    containsSemiMixedTypes: text || blank ? undefined : "0",
    containsNonDate: date && !number && !text ? "0" : undefined,
    containsDate: date ? "1" : undefined,
    containsString: text ? undefined : "0",
    containsBlank: blank ? "1" : undefined,
    containsMixedTypes: kinds.size > 1 ? "1" : undefined,
    containsNumber: number ? "1" : undefined,
    containsInteger: number && integer ? "1" : undefined,
    minValue: number ? min : undefined,
    maxValue: number ? max : undefined,
    minDate: date ? serialToIso(min) : undefined,
    maxDate: date ? serialToIso(max) : undefined,
    longText: long ? "1" : undefined,
  })
}

/** The x14 extension of a cache definition, which gives the cache the id PivotTable slicers name it by. */
const X15_FILTER = "{0605FD5F-26C8-4aeb-8148-2DB25E43C511}"
const NS_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
export const X14_CACHE = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}"

export function pivotCacheXml(cache: PivotCache, recordsRelId: string, numFmtId: (code?: string) => number): string {
  const src =
    cache.source.kind === "name"
      ? `<worksheetSource${attrText({ name: cache.source.name })}/>`
      : `<worksheetSource${attrText({ ref: cache.source.ref, sheet: cache.source.sheet })}/>`
  const fields = cache.fields.map((f, i) => {
    const dates = f.dates ?? false
    const attrs = attrText({ name: f.name, numFmtId: numFmtId(f.numFmt) })
    if (f.formula !== undefined) return `<cacheField${attrs}${attrText({ formula: f.formula, databaseField: 0 })}/>`
    const group = fieldGroupXml(cache, i)
    if (!inRecords(f, i)) return `<cacheField${attrs}${attrText({ databaseField: 0 })}>${group}</cacheField>`
    if (!f.items) {
      const values = (function* () {
        for (const r of cache.records) yield r[i] ?? null
      })()
      return `<cacheField${attrs}><sharedItems${sharedItemsAttrs(values, dates)}/>${group}</cacheField>`
    }
    const items = f.items.map((v) => valueXml(v, dates)).join("")
    return (
      `<cacheField${attrs}><sharedItems${sharedItemsAttrs(f.items, dates)} count="${f.items.length}">` +
      `${items}</sharedItems>${group}</cacheField>`
    )
  })
  return (
    `<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}"` +
    attrText({
      "r:id": recordsRelId,
      refreshOnLoad: cache.refreshOnLoad || undefined,
      createdVersion: 8,
      refreshedVersion: 8,
      minRefreshableVersion: 3,
      recordCount: cache.records.length,
    }) +
    `><cacheSource type="worksheet">${src}</cacheSource>` +
    `<cacheFields count="${fields.length}">${fields.join("")}</cacheFields>` +
    (cache.slicerId !== undefined
      ? `<extLst><ext uri="${X14_CACHE}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${cache.slicerId}"/></ext></extLst>`
      : "") +
    `</pivotCacheDefinition>`
  )
}

/** The records: shared items by index, other values inline. */
export function pivotRecordsXml(cache: PivotCache): string {
  const index = cache.fields.map((f) => {
    if (!f.items) return null
    const m = new Map<string, number>()
    f.items.forEach((v, x) => m.has(valueKey(v)) || m.set(valueKey(v), x))
    return m
  })
  const rows: string[] = []
  for (const rec of cache.records) {
    let row = "<r>"
    cache.fields.forEach((f, i) => {
      if (!inRecords(f, i)) return
      const v = rec[i] ?? null
      const x = index[i]?.get(valueKey(v))
      row += x !== undefined ? `<x v="${x}"/>` : valueXml(v, f.dates ?? false)
    })
    rows.push(row + "</r>")
  }
  return (
    `<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_R}" count="${cache.records.length}">` +
    `${rows.join("")}</pivotCacheRecords>`
  )
}

const AXIS_ATTR = { row: "axisRow", col: "axisCol", page: "axisPage" } as const
const ITEM_TYPE = (t: PivotAxisItem["t"]) =>
  t === "data" ? undefined : t === "default" || t === "grand" || t === "blank" ? t : FN_ITEM[t]

function axisItemsXml(local: string, items: PivotAxisItem[]): string {
  const body = items
    .map((i) => {
      const xs = i.x.map((v) => (v ? `<x v="${v}"/>` : "<x/>")).join("")
      const attrs = attrText({ t: ITEM_TYPE(i.t), r: i.r || undefined, i: i.i || undefined })
      return xs ? `<i${attrs}>${xs}</i>` : `<i${attrs}/>`
    })
    .join("")
  return items.length ? `<${local} count="${items.length}">${body}</${local}>` : ""
}

/** The order of pivotTableDefinition's children. */
const TABLE_ORDER = [
  "location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "dataFields", "formats",
  "conditionalFormats", "chartFormats", "pivotHierarchies", "pivotTableStyleInfo", "filters", "rowHierarchiesUsage",
  "colHierarchiesUsage", "extLst",
] // prettier-ignore

/** A PivotTable definition, with the rows and columns the engine lays out from the cache. */
export function pivotTableXml(
  table: PivotTable,
  cache: PivotCache,
  numFmtId: (code?: string) => number,
  kept: readonly Kept[] = [],
): string {
  const layout = pivotLayout(table, cache)
  /** A base item as the file keeps it: its position in the base field's written items. */
  const baseItemAttr = (v: PivotTable["values"][number]): number => {
    if (v.baseItem === "previous") return BASE_PREVIOUS
    if (v.baseItem === "next") return BASE_NEXT
    if (v.baseItem === undefined) return 0
    const f = v.baseField ?? 0
    const field = table.fields[f]
    if (!field || (field.axis === null && !field.items.length)) return v.baseItem
    const pos = fullItems(field, sharedCount(cache, f)).findIndex((it) => it.x === v.baseItem)
    return Math.max(pos, 0)
  }
  const fields = table.fields.map((f, i) => {
    const custom = f.subtotals.filter((s) => s !== "default") as PivotFn[]
    const hasDefault = f.subtotals.includes("default")
    const attrs: Record<string, string | number | undefined> = {
      name: f.name,
      axis: f.axis ? AXIS_ATTR[f.axis] : undefined,
      dataField: f.dataField ? 1 : undefined,
      // A calculated field can only be a value.
      ...(cache.fields[i]?.formula !== undefined ? { dragToRow: 0, dragToCol: 0, dragToPage: 0 } : {}),
      compact: f.compact ? undefined : 0,
      numFmtId: f.numFmt ? numFmtId(f.numFmt) : undefined,
      outline: f.outline ? undefined : 0,
      subtotalTop: f.subtotalTop ? undefined : 0,
      multipleItemSelectionAllowed: f.multipleItems ? 1 : undefined,
      measureFilter: table.filters?.some((x) => x.field === i && (x.kind === "value" || x.kind === "top"))
        ? 1
        : undefined,
      insertBlankRow: f.insertBlankRow ? 1 : undefined,
      showAll: f.showAll ? undefined : 0,
      sortType: f.sort === "manual" ? undefined : f.sort,
      defaultSubtotal: hasDefault || custom.length ? undefined : 0,
    }
    for (const [fn, attr] of FN_ATTR) if (custom.includes(fn)) attrs[attr] = 1
    let inner = ""
    const listed = f.axis !== null || f.items.length > 0
    if (listed) {
      // Every shared item, once: the field's own order, then any it doesn't list yet.
      const items = fullItems(f, sharedCount(cache, i))
      const xml = items.map(
        (it) =>
          `<item${attrText({ n: it.name, h: it.hidden ? 1 : undefined, sd: it.collapsed ? 0 : undefined, x: it.x })}/>`,
      )
      if (custom.length) for (const fn of custom) xml.push(`<item t="${FN_ITEM[fn]}"/>`)
      else if (hasDefault) xml.push(`<item t="default"/>`)
      inner += `<items count="${xml.length}">${xml.join("")}</items>`
    }
    if (f.repeatLabels)
      inner +=
        `<extLst><ext uri="${X14_FIELD}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
        `<x14:pivotField fillDownLabels="1"/></ext></extLst>`
    return inner ? `<pivotField${attrText(attrs)}>${inner}</pivotField>` : `<pivotField${attrText(attrs)}/>`
  })
  const fieldList = (local: string, list: number[]) =>
    list.length ? `<${local} count="${list.length}">${list.map((x) => `<field x="${x}"/>`).join("")}</${local}>` : ""
  const generated: Record<string, string> = {
    location: `<location${attrText({
      ref: rangeName(layout.range),
      firstHeaderRow: layout.firstHeaderRow,
      firstDataRow: layout.firstDataRow,
      firstDataCol: layout.firstDataCol,
      rowPageCount: table.pages.length || undefined,
      colPageCount: table.pages.length ? 1 : undefined,
    })}/>`,
    pivotFields: `<pivotFields count="${fields.length}">${fields.join("")}</pivotFields>`,
    rowFields: fieldList("rowFields", table.rows),
    rowItems: axisItemsXml("rowItems", layout.rowItems),
    colFields: fieldList("colFields", table.cols),
    colItems: axisItemsXml("colItems", layout.colItems),
    pageFields: table.pages.length
      ? `<pageFields count="${table.pages.length}">${table.pages
          .map((p) => `<pageField${attrText({ fld: p.field, item: p.item ?? undefined })} hier="-1"/>`)
          .join("")}</pageFields>`
      : "",
    dataFields: table.values.length
      ? `<dataFields count="${table.values.length}">${table.values
          .map(
            (v) =>
              `<dataField${attrText({
                name: v.name,
                fld: v.field,
                subtotal: v.fn === "sum" ? undefined : v.fn,
                showDataAs: v.showAs && SHOW_DATA_AS.has(v.showAs) ? v.showAs : undefined,
                baseField: v.baseField ?? 0,
                baseItem: baseItemAttr(v),
                numFmtId: v.numFmt ? numFmtId(v.numFmt) : undefined,
              })}${
                v.showAs && !SHOW_DATA_AS.has(v.showAs)
                  ? `><extLst><ext uri="${X14_DATA}" xmlns:x14="${NS_X14}"><x14:dataField pivotShowAs="${v.showAs}"/></ext></extLst></dataField>`
                  : "/>"
              }`,
          )
          .join("")}</dataFields>`
      : "",
    filters: table.filters?.length
      ? `<filters count="${table.filters.length}">${table.filters.map(pivotFilterXml).join("")}</filters>`
      : "",
    pivotTableStyleInfo: `<pivotTableStyleInfo${attrText({
      name: table.style.name ?? undefined,
      showRowHeaders: table.style.rowHeaders ? 1 : 0,
      showColHeaders: table.style.colHeaders ? 1 : 0,
      showRowStripes: table.style.rowStripes ? 1 : 0,
      showColStripes: table.style.colStripes ? 1 : 0,
      showLastColumn: table.style.lastColumn ? 1 : 0,
    })}/>`,
  }
  const c = table.captions
  const attrs = attrText({
    name: table.name,
    cacheId: table.cacheId,
    dataOnRows: table.rows.includes(VALUES_FIELD) ? 1 : undefined,
    applyNumberFormats: 0,
    applyBorderFormats: 0,
    applyFontFormats: 0,
    applyPatternFormats: 0,
    applyAlignmentFormats: 0,
    applyWidthHeightFormats: 1,
    dataCaption: c.data,
    grandTotalCaption: c.grandTotal,
    showError: table.showError ? 1 : undefined,
    errorCaption: c.error,
    showMissing: table.showMissing ? undefined : 0,
    missingCaption: c.missing,
    updatedVersion: 8,
    minRefreshableVersion: 3,
    useAutoFormatting: table.autoFit === false ? undefined : 1,
    rowGrandTotals: table.rowGrandTotals ? undefined : 0,
    colGrandTotals: table.colGrandTotals ? undefined : 0,
    itemPrintTitles: 1,
    createdVersion: 8,
    indent: 0,
    compact: table.compact ? undefined : 0,
    compactData: table.compactData ? undefined : 0,
    outline: table.outline ? 1 : undefined,
    outlineData: table.outline ? 1 : undefined,
    enableDrill: table.enableDrill === false ? 0 : undefined,
    showDrill: table.showDrill === false ? 0 : undefined,
    multipleFieldFilters: 0,
    rowHeaderCaption: c.rowHeader,
    colHeaderCaption: c.colHeader,
  })
  // A PivotTable made in Cells says Excel 365's "no Values row" the way Excel does; a read one keeps its extLst.
  if (table.hideValuesRow && !kept.some((k) => k.local === "extLst"))
    generated.extLst =
      `<extLst><ext uri="${X14_TABLE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
      `<x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/>` +
      `</ext></extLst>`
  const body = TABLE_ORDER.map((local) => generated[local] ?? kept.find((k) => k.local === local)?.xml ?? "").join("")
  return `<pivotTableDefinition xmlns="${NS_MAIN}"${attrs}>${body}</pivotTableDefinition>`
}
