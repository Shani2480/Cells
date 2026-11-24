import { CALC_ITEM_FNS, calcItemValue } from "./calc-items"
import type { Range } from "../address"
import { wildcardRegExp } from "../filter"
import { errorValue, type CellError, type CellValue } from "../values"
import {
  PIVOT_FN_LABEL,
  VALUES_FIELD,
  type PivotCache,
  type PivotDataField,
  type PivotFilter,
  type PivotFn,
  type PivotSubtotal,
  type PivotTable,
  type PivotValue,
} from "./model"
import { showAsNumFmt, showValueAs, type ShowAsContext } from "./show-as"
import { calcFieldInputs, calcFieldValue } from "./calculated"
import { fieldShared, recordXs } from "./group"

/**
 * The PivotTable engine: the cells a PivotTable shows, worked out from its cache and settings the way Excel lays them
 * out (checked cell by cell against Excel's own files, packages/io/test/fixtures/excel/pivot-*.xlsx).
 */

/** What a cell of the PivotTable is, for its look (PivotTable styles) and for clicks on it. */
export type PivotCellRole =
  | "pageLabel"
  | "pageValue"
  | "corner"
  | "colHeader"
  | "rowHeader"
  | "dataHeader"
  | "colItem"
  | "rowItem"
  | "subtotalLabel"
  | "grandLabel"
  | "value"
  | "subtotalValue"
  | "grandValue"

export type PivotCell = {
  row: number
  col: number
  /** Undefined: an empty cell of the table (it still takes the table's look). */
  value?: CellValue
  role: PivotCellRole
  /** Compact form: how far the label is indented (levels sharing its column before it). */
  indent?: number
  numFmt?: string
  /** A row or column label: the field and cache item it stands for, its axis and level on it. */
  item?: PivotCellItem
  /** A value cell: its line on the row and column axes (what Show Details looks up). */
  lines?: [number, number]
  /** The +/− button the label shows (an outer item with a field inside it, PivotTable Options › Show expand/collapse buttons). */
  button?: "collapse" | "expand"
}

export type PivotCellItem = { field: number; x: number; axis: "row" | "col"; level: number }

/** One line of the row or column area, as Excel's rowItems / colItems keep it. */
export type PivotAxisItem = {
  /** "data" for an item's own line; a subtotal's function; "grand"; "blank". */
  t: "data" | "blank" | "grand" | PivotSubtotal
  /** Leading members repeated from the line before. */
  r: number
  /** The value field (when the Values field is on this axis). */
  i: number
  /** Item indexes from level `r` on (the Values level holds the value field's index). */
  x: number[]
}

export type PivotLayout = {
  /** The table's cells (Excel's location ref), without the report filters. */
  range: Range
  /** The report filters' cells, above the table. */
  pages: Range | null
  firstHeaderRow: number
  firstDataRow: number
  firstDataCol: number
  /** The label column (from the table's left) of each row level; compact levels share their parent's. */
  rowLabelCols: number[]
  cells: PivotCell[]
  rowItems: PivotAxisItem[]
  colItems: PivotAxisItem[]
}

/** A line of an axis while it is worked out. */
type Line = {
  t: PivotAxisItem["t"]
  /** Item positions per level (the Values level holds the value field's index). */
  path: number[]
  /** The value field the line is for; -1 when it isn't settled on this axis (a level above the Values field). */
  data: number
  /** The line shows values. */
  values: boolean
}

// ── Values ──

/** Excel's sort order for items: numbers, text, logical values, errors, blanks. */
function rank(v: PivotValue): number {
  if (v === null || v === "") return 4
  if (typeof v === "number") return 0
  if (typeof v === "string") return 1
  if (typeof v === "boolean") return 2
  return 3
}

export function comparePivotValues(a: PivotValue, b: PivotValue): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === "number" && typeof b === "number") return a - b
  if (typeof a === "string" && typeof b === "string") return a.localeCompare(b, undefined, { sensitivity: "base" })
  if (typeof a === "boolean" && typeof b === "boolean") return Number(a) - Number(b)
  return 0
}

/** Running figures for one group of records, enough for every summary function. */
type Stats = {
  n: number
  nums: number
  sum: number
  prod: number
  min: number
  max: number
  mean: number
  m2: number
  error: CellError | null
}

const newStats = (): Stats => ({
  n: 0,
  nums: 0,
  sum: 0,
  prod: 1,
  min: Infinity,
  max: -Infinity,
  mean: 0,
  m2: 0,
  error: null,
})

function addValue(s: Stats, v: PivotValue) {
  if (v === null || v === "") return
  s.n++
  if (typeof v === "number") {
    s.nums++
    s.sum += v
    s.prod *= v
    if (v < s.min) s.min = v
    if (v > s.max) s.max = v
    // Welford's running variance.
    const d = v - s.mean
    s.mean += d / s.nums
    s.m2 += d * (v - s.mean)
  } else if (typeof v === "object" && !s.error) s.error = v
}

const DIV0 = () => errorValue("#DIV/0!")

/** A group's summary (undefined: the cell stays empty, as Excel leaves groups with no numbers). */
export function summarize(s: Stats, fn: PivotFn): CellValue | undefined {
  if (fn === "count") return s.n || undefined
  if (fn === "countNums") return s.nums || undefined
  if (s.error) return s.error
  if (!s.nums) return undefined
  switch (fn) {
    case "sum":
      return s.sum
    case "average":
      return s.sum / s.nums
    case "max":
      return s.max
    case "min":
      return s.min
    case "product":
      return s.prod
    case "stdDev":
      return s.nums < 2 ? DIV0() : Math.sqrt(s.m2 / (s.nums - 1))
    case "stdDevp":
      return Math.sqrt(s.m2 / s.nums)
    case "var":
      return s.nums < 2 ? DIV0() : s.m2 / (s.nums - 1)
    case "varp":
      return s.m2 / s.nums
  }
}

// ── Items ──

/** A field's items as the engine uses them: their values, display order, and where each record falls. */
type FieldItems = {
  /** The value of each item position. */
  values: PivotValue[]
  /** The cache's shared item of each position. */
  xs: number[]
  /** Item positions in the order shown (hidden ones left out). */
  order: number[]
  hidden: Set<number>
  collapsed: Set<number>
  names: (string | undefined)[]
  /** The item position of each record. */
  ofRecord: Int32Array
  /** Calculated items: item position → index in the cache's `calcItems`. */
  calc: Map<number, number>
}

function fieldItems(table: PivotTable, cache: PivotCache, f: number): FieldItems {
  const field = table.fields[f]
  const cacheField = cache.fields[f]
  // The cache's shared items (a grouped field's group items), or the column's distinct values when it keeps none.
  const shared = fieldShared(cache, f)
  // Item positions follow the field's items; shared items it doesn't list yet go after them.
  const posOfShared = new Map<number, number>()
  const xs: number[] = []
  const hidden = new Set<number>()
  const collapsed = new Set<number>()
  const names: (string | undefined)[] = []
  for (const item of field?.items ?? []) {
    if (item.x < 0 || item.x >= shared.length || posOfShared.has(item.x)) continue
    const pos = xs.length
    posOfShared.set(item.x, pos)
    xs.push(item.x)
    if (item.hidden) hidden.add(pos)
    if (item.collapsed) collapsed.add(pos)
    names.push(item.name)
  }
  for (let x = 0; x < shared.length; x++)
    if (!posOfShared.has(x)) {
      posOfShared.set(x, xs.length)
      xs.push(x)
      names.push(undefined)
    }
  const values = xs.map((x) => shared[x])
  const xsOfRecord = recordXs(cache, f)
  const ofRecord = new Int32Array(xsOfRecord.length)
  for (let i = 0; i < xsOfRecord.length; i++) ofRecord[i] = posOfShared.get(xsOfRecord[i]) ?? -1
  let order = xs.map((_, pos) => pos).filter((pos) => !hidden.has(pos))
  if (field?.sort === "ascending" || field?.sort === "descending") {
    const sign = field.sort === "ascending" ? 1 : -1
    // Ranges and date parts sort in their own order (Jan before Feb), not by their labels.
    const ranged = cacheField.group?.kind === "range"
    order = order.sort((a, b) => {
      if (ranged) return sign * (xs[a] - xs[b])
      const va = names[a] ?? values[a]
      const vb = names[b] ?? values[b]
      // Blanks stay last either way.
      const blanks = Number(rank(va) === 4) - Number(rank(vb) === 4)
      return blanks || sign * comparePivotValues(va, vb)
    })
  }
  const calc = new Map<number, number>()
  cache.calcItems?.forEach((ci, k) => {
    const pos = ci.field === f ? posOfShared.get(ci.x) : undefined
    if (pos !== undefined) calc.set(pos, k)
  })
  return { values, xs, order, hidden, collapsed, names, ofRecord, calc }
}

/** An item's label: its typed name, or its value ("(blank)" for empty). */
function itemLabel(items: FieldItems, pos: number): CellValue {
  const name = items.names[pos]
  if (name !== undefined) return name
  const v = items.values[pos]
  return v === null || v === "" ? "(blank)" : v
}

const labelText = (v: CellValue): string =>
  typeof v === "object" ? v.error : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)

/** A field's name as the PivotTable shows it. */
export function pivotFieldName(table: PivotTable, cache: PivotCache, f: number): string {
  if (f === VALUES_FIELD) return table.captions.data
  return table.fields[f]?.name ?? cache.fields[f]?.name ?? ""
}

// ── Axes ──

/** A group of records under one path of items (`records` kept only when an item on the axis sorts by value). */
type Node = { children: Map<number, Node>; records?: number[] }

/** The value field (and direction) that orders field `f`'s items, or null when they follow their labels. */
function valueSortOf(table: PivotTable, f: number): { df: PivotDataField; sign: number } | null {
  const field = table.fields[f]
  if (!field || field.sortBy === undefined || field.sort === "manual") return null
  const df = table.values[field.sortBy]
  return df ? { df, sign: field.sort === "ascending" ? 1 : -1 } : null
}

/**
 * Items in the order of their totals of a value field under `node` (Excel's autoSort by a value): smallest or biggest
 * first; items with no number (no data, or text summed) go last, keeping their label order, as do ties.
 */
function sortedByValue(
  order: number[],
  node: Node,
  { df, sign }: { df: PivotDataField; sign: number },
  cache: PivotCache,
): number[] {
  const inputs = calcFieldInputs(cache, df.field)
  const total = (pos: number): number | null => {
    const records = node.children.get(pos)?.records
    if (!records) return null
    const stats = new Map(inputs.map((f) => [f, newStats()]))
    for (const r of records) for (const f of inputs) addValue(stats.get(f)!, cache.records[r][f] ?? null)
    const v = cache.fields[df.field]?.formula
      ? calcFieldValue(cache, df.field, (g) => stats.get(g)?.sum ?? 0)
      : summarize(stats.get(df.field)!, df.fn)
    return typeof v === "number" ? v : null
  }
  const totals = new Map(order.map((pos) => [pos, total(pos)]))
  return order
    .map((pos, i) => ({ pos, i, v: totals.get(pos)! }))
    .sort((a, b) => {
      if (a.v === null || b.v === null) return Number(a.v === null) - Number(b.v === null) || a.i - b.i
      return sign * (a.v - b.v) || a.i - b.i
    })
    .map((x) => x.pos)
}

/** The items under any of `nodes` (what a calculated item holds inside it: its siblings' inner items). */
function mergedNode(nodes: Node[]): Node {
  const children = new Map<number, Node[]>()
  for (const n of nodes) for (const [pos, child] of n.children) children.set(pos, [...(children.get(pos) ?? []), child])
  return { children: new Map([...children].map(([pos, list]) => [pos, mergedNode(list)])) }
}

function axisLines(
  cache: PivotCache,
  table: PivotTable,
  fields: number[],
  items: Map<number, FieldItems>,
  records: number[],
  rows: boolean,
): Line[] {
  const valueCount = table.values.length
  const dataAt = fields.indexOf(VALUES_FIELD)
  const real = fields.filter((f) => f !== VALUES_FIELD)
  // The tree of item paths that have data.
  const byValue = real.some((f) => valueSortOf(table, f) !== null)
  const root: Node = { children: new Map() }
  for (const r of records) {
    let node = root
    for (const f of real) {
      const pos = items.get(f)!.ofRecord[r]
      let next = node.children.get(pos)
      if (!next) node.children.set(pos, (next = { children: new Map() }))
      node = next
      if (byValue) (node.records ??= []).push(r)
    }
  }
  const lines: Line[] = []
  const dataOf = (path: number[]) => (dataAt >= 0 ? (path.length > dataAt ? path[dataAt] : -1) : 0)

  const walk = (level: number, path: number[], node: Node) => {
    const f = fields[level]
    const last = level === fields.length - 1
    let children: { pos: number; node: Node }[]
    if (f === VALUES_FIELD) children = table.values.map((_, i) => ({ pos: i, node }))
    else {
      const fi = items.get(f)!
      const field = table.fields[f]
      // A calculated item shows under every item, whether or not the items it names have data there.
      let shown = field.showAll ? fi.order : fi.order.filter((pos) => node.children.has(pos) || fi.calc.has(pos))
      const by = valueSortOf(table, f)
      if (by) shown = sortedByValue(shown, node, by, cache)
      children = shown.map((pos) => ({
        pos,
        node:
          node.children.get(pos) ??
          (fi.calc.has(pos) ? mergedNode([...node.children.values()]) : { children: new Map() }),
      }))
    }
    for (const { pos, node: child } of children) {
      const p = [...path, pos]
      const field = f === VALUES_FIELD ? null : table.fields[f]
      const collapsed = !last && f !== VALUES_FIELD && items.get(f)!.collapsed.has(pos)
      if (last || collapsed) {
        lines.push({ t: "data", path: p, data: dataOf(p), values: dataOf(p) >= 0 })
        continue
      }
      const tabular = field ? !field.outline : false
      // A field with only the Values field inside it has no subtotals: the value lines are its totals.
      const onlyValues = level + 1 === dataAt && dataAt === fields.length - 1
      const subs = field && !onlyValues ? field.subtotals : []
      const top = rows && !tabular && subs.length === 1 && subs[0] === "default" && field!.subtotalTop
      // Rows give an outer item its own line; columns put its label over its inner items instead.
      if (rows && !tabular) lines.push({ t: "data", path: p, data: dataOf(p), values: top && dataOf(p) >= 0 })
      walk(level + 1, p, child)
      if (!top)
        for (const s of subs) {
          // Under a Values field still to come, a subtotal line per value field.
          if (dataOf(p) < 0) for (let i = 0; i < valueCount; i++) lines.push({ t: s, path: p, data: i, values: true })
          else lines.push({ t: s, path: p, data: dataOf(p), values: true })
        }
      if (rows && field?.insertBlankRow) lines.push({ t: "blank", path: p, data: -1, values: false })
    }
  }

  if (!fields.length) return [{ t: "data", path: [], data: valueCount ? 0 : -1, values: valueCount > 0 }]
  walk(0, [], root)
  const grand = rows ? table.rowGrandTotals : table.colGrandTotals
  if (grand && real.length) {
    if (dataAt >= 0) for (let i = 0; i < valueCount; i++) lines.push({ t: "grand", path: [0], data: i, values: true })
    else lines.push({ t: "grand", path: [0], data: 0, values: valueCount > 0 })
  }
  return lines
}

/** Excel's rowItems / colItems for the lines of an axis. */
function axisItems(lines: Line[], hasData: boolean): PivotAxisItem[] {
  if (lines.length === 1 && !lines[0].path.length) return [{ t: "data", r: 0, i: 0, x: [] }]
  let prev: number[] = []
  return lines.map((line) => {
    let r = 0
    // Subtotal and grand total lines spell out their whole path; item and blank lines skip what they repeat.
    if (line.t === "data" || line.t === "blank")
      while (r < line.path.length - 1 && r < prev.length && prev[r] === line.path[r]) r++
    prev = line.path
    return { t: line.t, r, i: hasData && line.data > 0 ? line.data : 0, x: line.path.slice(r) }
  })
}

/** A subtotal line's own function (null: the value field's). */
const ownFn = (t: PivotAxisItem["t"]): PivotFn | null =>
  t === "data" || t === "grand" || t === "default" || t === "blank" ? null : t

// ── The layout ──

/** The first half of the layout: the fields' items, the records that count and each axis' lines. */
function prepare(table: PivotTable, cache: PivotCache) {
  const rowFields = table.rows
  const colFields = table.cols
  const items = new Map<number, FieldItems>()
  for (const f of [...rowFields, ...colFields, ...table.pages.map((p) => p.field)])
    if (f !== VALUES_FIELD && !items.has(f) && cache.fields[f]) items.set(f, fieldItems(table, cache, f))

  // A field on no axis with items hidden: a slicer's (Excel keeps a slicer's picks as hidden items of each
  // PivotTable it filters, whether the field is in the PivotTable or not).
  const sliced: FieldItems[] = []
  table.fields.forEach((field, f) => {
    if (!field?.items.some((i) => i.hidden) || items.has(f) || !cache.fields[f]) return
    sliced.push(fieldItems(table, cache, f))
  })

  // A timeline's dates, on whatever field (whole days: the last day counts to its end).
  const dates = (table.filters ?? []).flatMap((x) => (x.kind === "date" && cache.fields[x.field] ? [x] : []))

  // The records that count: report filters' picks, no hidden item on a row or column field, a slicer's picks and a
  // timeline's dates.
  const records: number[] = []
  next: for (let r = 0; r < cache.records.length; r++) {
    for (const fi of sliced) if (fi.hidden.has(fi.ofRecord[r])) continue next
    for (const d of dates) {
      const v = cache.records[r][d.field]
      if (typeof v !== "number" || v < d.value || v >= Math.floor(d.value2) + 1) continue next
    }
    for (const page of table.pages) {
      const fi = items.get(page.field)
      if (!fi) continue
      const pos = fi.ofRecord[r]
      if (page.item !== null ? pos !== page.item : fi.hidden.has(pos)) continue next
    }
    for (const f of [...rowFields, ...colFields]) {
      if (f === VALUES_FIELD) continue
      if (items.get(f)!.hidden.has(items.get(f)!.ofRecord[r])) continue next
    }
    records.push(r)
  }
  const filtered = applyFilters(table, cache, items, records)

  const rowLines = axisLines(cache, table, rowFields, items, filtered.rows, true)
  const colLines = axisLines(cache, table, colFields, items, filtered.cols, false)
  return { rowFields, colFields, items, filtered, rowLines, colLines }
}

/** One line of an axis, for GETPIVOTDATA: its kind, the value field it is for and its items on the real fields. */
export type PivotAxisLine = { t: PivotAxisItem["t"]; data: number; values: boolean; items: number[] }

/**
 * A PivotTable's axes as GETPIVOTDATA reads them: each axis' real fields and lines, and a way to find an item of a
 * field by its value or label (`match` compares a label text, ignoring case, or a number).
 */
export function pivotAxes(table: PivotTable, cache: PivotCache) {
  const { rowFields, colFields, items, rowLines, colLines } = prepare(table, cache)
  const lines = (list: Line[], fields: number[]): PivotAxisLine[] =>
    list.map((l) => ({
      t: l.t,
      data: l.data,
      values: l.values,
      items: l.t === "grand" ? [] : l.path.filter((_, j) => fields[j] !== VALUES_FIELD),
    }))
  /** The item position of field `f` that `want` names (shown items only), or -1. */
  const itemOf = (f: number, want: string | number | boolean): number => {
    const fi = items.get(f)
    if (!fi) return -1
    // A report filter's picked item counts even when the field has other items hidden.
    const candidates = table.pages.some((p) => p.field === f) ? fi.xs.map((_, pos) => pos) : fi.order
    return (
      candidates.find((pos) => {
        const name = fi.names[pos]
        const v = fi.values[pos]
        if (name === undefined && typeof want === "number" && typeof v === "number") return v === want
        // A month or quarter by its number (1 is Jan), as Excel takes it.
        const by = (cache.fields[f]?.group as { by?: string } | undefined)?.by
        if (typeof want === "number" && (by === "months" || by === "quarters")) return fi.xs[pos] === want
        const label = labelText(itemLabel(fi, pos)).toLowerCase()
        return label === String(want).toLowerCase()
      }) ?? -1
    )
  }
  return {
    rowFields: rowFields.filter((f) => f !== VALUES_FIELD),
    colFields: colFields.filter((f) => f !== VALUES_FIELD),
    rowLines: lines(rowLines, rowFields),
    colLines: lines(colLines, colFields),
    itemOf,
    /** An item's typed name, or its value. */
    itemValue: (f: number, pos: number): PivotValue => {
      const fi = items.get(f)
      return fi ? (fi.names[pos] ?? fi.values[pos] ?? null) : null
    },
  }
}

/**
 * Show Details (double-clicking a value): the cache records a value cell of the layout sums up, in the cache's order;
 * null for a cell that isn't a value.
 */
export function pivotDetailRecords(table: PivotTable, cache: PivotCache, cell: PivotCell): number[] | null {
  if (!cell.lines) return null
  const { rowFields, colFields, items, filtered, rowLines, colLines } = prepare(table, cache)
  const rowLine = rowLines[cell.lines[0]]
  const colLine = colLines[cell.lines[1]]
  if (!rowLine || !colLine) return null
  const conditions: [FieldItems, number][] = []
  const add = (line: Line, fields: number[]) => {
    if (line.t === "grand") return
    line.path.forEach((pos, j) => {
      if (fields[j] !== VALUES_FIELD) conditions.push([items.get(fields[j])!, pos])
    })
  }
  add(rowLine, rowFields)
  add(colLine, colFields)
  // A calculated item's value has no records of its own (Excel won't show details there).
  if (conditions.some(([fi, pos]) => fi.calc.has(pos))) return null
  return filtered.both.filter((r) => conditions.every(([fi, pos]) => fi.ofRecord[r] === pos))
}

export function pivotLayout(table: PivotTable, cache: PivotCache): PivotLayout {
  const { rowFields, colFields, items, filtered, rowLines, colLines } = prepare(table, cache)

  // Label columns: a compact field shares its column with the field after it.
  const compactOf = (f: number) =>
    f === VALUES_FIELD ? table.compactData : table.fields[f].compact && table.fields[f].outline
  const labelCol: number[] = []
  rowFields.forEach((_f, j) => labelCol.push(j === 0 ? 0 : labelCol[j - 1] + (compactOf(rowFields[j - 1]) ? 0 : 1)))
  const labelCols = rowFields.length ? labelCol[labelCol.length - 1] + 1 : 0
  const n = colFields.length
  // Excel 365 leaves out the row above the value fields' names when they are the only column field (the Values row).
  const valuesRowHidden = Boolean(table.hideValuesRow) && n === 1 && colFields[0] === VALUES_FIELD
  const headerRows = n ? n + (valuesRowHidden ? 0 : 1) : 1
  const top = table.at.row
  const left = table.at.col
  const cells: PivotCell[] = []
  const put = (row: number, col: number, role: PivotCellRole, value?: CellValue, extra?: Partial<PivotCell>) =>
    cells.push({ row: top + row, col: left + col, role, ...(value !== undefined ? { value } : {}), ...extra })
  const dataName = (i: number) => table.values[i]?.name ?? ""
  const label = (f: number, pos: number): CellValue =>
    f === VALUES_FIELD ? dataName(pos) : itemLabel(items.get(f)!, pos)
  const subLabel = (line: Line, f: number, dataBelow: boolean) => {
    const own = labelText(label(f, line.path[line.path.length - 1]))
    const fn = line.t === "default" ? null : PIVOT_FN_LABEL[line.t as PivotFn]
    if (dataBelow && table.values.length > 1) return `${own}${fn ? ` ${fn}` : ""} ${dataName(line.data)}`
    return `${own} ${fn ?? "Total"}`
  }
  const grandLabel = (line: Line, dataOnAxis: boolean) =>
    dataOnAxis && table.values.length > 1
      ? `Total ${dataName(line.data)}`
      : (table.captions.grandTotal ?? "Grand Total")

  /** What a label cell of field `f` at `level` stands for, and its button. */
  const itemOf = (axis: "row" | "col", fields: number[], level: number, pos: number, first: boolean) => {
    const f = fields[level]
    if (f === VALUES_FIELD) return undefined
    const fi = items.get(f)!
    const extra: Partial<PivotCell> = { item: { field: f, x: fi.xs[pos], axis, level } }
    // A number item shows in its field's format (a date column's, when the field has none of its own).
    const fmt = table.fields[f]?.numFmt ?? (cache.fields[f]?.dates ? cache.fields[f].numFmt : undefined)
    if (fmt && fi.names[pos] === undefined && typeof fi.values[pos] === "number") extra.numFmt = fmt
    if (first && table.showDrill !== false && fields.slice(level + 1).some((g) => g !== VALUES_FIELD))
      extra.button = fi.collapsed.has(pos) ? "expand" : "collapse"
    return extra
  }

  // ── The header rows ──
  const dataInRows = rowFields.includes(VALUES_FIELD)
  const dataInCols = colFields.includes(VALUES_FIELD)
  if (n && !valuesRowHidden) {
    // With no row fields the column header takes the corner's cell.
    if (!labelCols) {
      /* no corner */
    } else if (table.values.length === 1 && !dataInRows && !dataInCols) put(0, 0, "corner", dataName(0))
    else put(0, 0, "corner")
    if (colFields.length === 1 && dataInCols) put(0, labelCols, "colHeader", table.captions.data)
    else if (table.showHeaders === false) put(0, labelCols, "colHeader")
    else if (compactOf(colFields[0])) put(0, labelCols, "colHeader", table.captions.colHeader ?? "Column Labels")
    else colFields.forEach((f, j) => put(0, labelCols + j, "colHeader", pivotFieldName(table, cache, f)))
  }
  const headerRow = headerRows - 1
  const rowHeaderCaption = table.captions.rowHeader ?? "Row Labels"
  for (let c = 0; c < labelCols; c++) {
    const j = labelCol.indexOf(c)
    const f = rowFields[j]
    const shared = labelCol.filter((x) => x === c).length > 1
    const caption = shared || compactOf(f) ? rowHeaderCaption : pivotFieldName(table, cache, f)
    put(headerRow, c, "rowHeader", table.showHeaders === false ? undefined : caption)
  }
  if (!n && table.values.length === 1) put(0, labelCols, "dataHeader", dataName(0))

  // Column labels: each level's item over the first column of its group.
  const colLevelRow = (j: number) => (valuesRowHidden ? 0 : 1) + j
  colLines.forEach((line, ci) => {
    const col = labelCols + ci
    if (!n) return
    if (line.t === "grand") return put(colLevelRow(0), col, "grandLabel", grandLabel(line, dataInCols))
    if (line.t !== "data") {
      const level = line.path.length - 1
      return put(
        colLevelRow(level),
        col,
        "subtotalLabel",
        subLabel(line, colFields[level], dataInCols && colFields.indexOf(VALUES_FIELD) > level),
      )
    }
    const prev = colLines
      .slice(0, ci)
      .reverse()
      .find((l) => l.t === "data")
    line.path.forEach((pos, j) => {
      const same =
        prev !== undefined && prev.path.length > j && prev.path.slice(0, j + 1).every((p, k) => p === line.path[k])
      if (!same) put(colLevelRow(j), col, "colItem", label(colFields[j], pos), itemOf("col", colFields, j, pos, true))
    })
  })

  // ── Rows ──
  const indentOf = (j: number) => labelCol.slice(0, j).filter((c) => c === labelCol[j]).length
  const statsFor = makeStats(table, cache, items, filtered.both)
  const realRowFields = rowFields.filter((f) => f !== VALUES_FIELD)
  const realColFields = colFields.filter((f) => f !== VALUES_FIELD)
  /** A value field's figure for row items `rows` × column items `cols` (a calculated field's formula on the sums). */
  const plain = (field: number, fn: PivotFn, rows: readonly number[], cols: readonly number[]) =>
    cache.fields[field]?.formula
      ? calcFieldValue(cache, field, (g) => statsFor.at(rows, cols, g).sum)
      : summarize(statsFor.at(rows, cols, field), fn)
  const measure = withCalcItems(cache, items, realRowFields, realColFields, filtered.both, plain)
  const showAsContext = (field: number, fn: PivotFn): ShowAsContext => ({
    rowFields: realRowFields,
    colFields: realColFields,
    order: (f) => items.get(f)?.order ?? [],
    posOf: (f, x) => {
      const pos = items.get(f)?.xs.indexOf(x) ?? -1
      return pos >= 0 ? pos : undefined
    },
    raw: (rows, cols) => measure(field, fn, rows, cols),
  })
  rowLines.forEach((line, ri) => {
    const row = headerRows + ri
    if (line.t === "blank") return
    if (line.t === "grand") put(row, 0, "grandLabel", grandLabel(line, dataInRows))
    else if (line.t !== "data") {
      const level = line.path.length - 1
      put(
        row,
        labelCol[level],
        "subtotalLabel",
        subLabel(line, rowFields[level], dataInRows && rowFields.indexOf(VALUES_FIELD) > level),
      )
    } else if (rowFields.length) {
      const prev = rowLines
        .slice(0, ri)
        .reverse()
        .find((l) => l.t === "data")
      const own = line.path.length - 1
      line.path.forEach((pos, j) => {
        const field = rowFields[j] === VALUES_FIELD ? null : table.fields[rowFields[j]]
        const started = !prev || prev.path.length <= j || prev.path.slice(0, j + 1).some((p, k) => p !== line.path[k])
        // An outer item's label sits on its own line, except in tabular form, where it goes on its first inner line.
        const show =
          j === own ||
          (labelCol[j] !== labelCol[own] && ((field && !field.outline && started) || (field?.repeatLabels ?? false)))
        if (show)
          put(row, labelCol[j], "rowItem", label(rowFields[j], pos), {
            ...(indentOf(j) ? { indent: indentOf(j) } : {}),
            ...itemOf("row", rowFields, j, pos, started || j === own),
          })
      })
    }
    // The values of the line.
    colLines.forEach((colLine, ci) => {
      const col = labelCols + ci
      const role: PivotCellRole =
        line.t === "grand" || colLine.t === "grand"
          ? "grandValue"
          : line.t !== "data" || colLine.t !== "data" || line.path.length < rowFields.length
            ? "subtotalValue"
            : "value"
      if (!line.values || !colLine.values) return put(row, col, role)
      const data = dataInRows ? line.data : dataInCols ? colLine.data : 0
      const df = table.values[data]
      if (!df) return put(row, col, role)
      const fn: PivotFn = ownFn(line.t) ?? ownFn(colLine.t) ?? df.fn
      const rowPath = statsFor.realPath(line, rowFields)
      const colPath = statsFor.realPath(colLine, colFields)
      let value = measure(df.field, fn, rowPath, colPath)
      if (df.showAs) value = showValueAs(df, showAsContext(df.field, df.fn), rowPath, colPath, value)
      if (value !== undefined && typeof value === "object" && table.showError && table.captions.error !== undefined)
        value = table.captions.error
      if (value === undefined && table.showMissing && table.captions.missing) value = table.captions.missing
      const numFmt = df.numFmt ?? showAsNumFmt(df.showAs)
      put(row, col, role, value, { lines: [ri, ci], ...(numFmt ? { numFmt } : {}) })
    })
  })

  // ── Report filters, above the table with a blank row between ──
  const pageCount = table.pages.length
  table.pages.forEach((page, i) => {
    const row = -(pageCount + 1) + i
    const fi = items.get(page.field)
    put(row, 0, "pageLabel", pivotFieldName(table, cache, page.field))
    let shown: CellValue = "(All)"
    if (fi && page.item !== null) shown = itemLabel(fi, page.item)
    else if (fi && fi.hidden.size) {
      const visible = fi.order
      shown = visible.length === 1 ? itemLabel(fi, visible[0]) : "(Multiple Items)"
    }
    put(row, 1, "pageValue", shown)
  })

  const width = labelCols + colLines.length
  const height = headerRows + rowLines.length
  if (!rowFields.length && !colFields.length && !table.values.length)
    // An empty PivotTable: Excel keeps 18 rows by 3 columns for its placeholder, and only report filters show.
    return {
      range: { r0: top, c0: left, r1: top + EMPTY_PIVOT.rows - 1, c1: left + EMPTY_PIVOT.cols - 1 },
      pages: pageCount ? { r0: top - pageCount - 1, c0: left, r1: top - 2, c1: left + 1 } : null,
      firstHeaderRow: 1,
      firstDataRow: 1,
      firstDataCol: 0,
      rowLabelCols: [],
      cells: cells.filter((c) => c.role === "pageLabel" || c.role === "pageValue"),
      rowItems: [],
      colItems: [],
    }
  return {
    range: { r0: top, c0: left, r1: top + height - 1, c1: left + Math.max(width, 1) - 1 },
    pages: pageCount ? { r0: top - pageCount - 1, c0: left, r1: top - 2, c1: left + 1 } : null,
    firstHeaderRow: valuesRowHidden ? 0 : 1,
    firstDataRow: headerRows,
    firstDataCol: labelCols,
    rowLabelCols: labelCol,
    cells,
    rowItems: axisItems(rowLines, dataInRows),
    colItems: axisItems(colLines, dataInCols),
  }
}

/** The area an empty PivotTable keeps for its placeholder ("To build a report, choose fields..."). */
export const EMPTY_PIVOT = { rows: 18, cols: 3 }

// ── Label Filters, Value Filters and Top 10 ──

/**
 * A label filter's test of one item's label. Labels compare as text ignoring case, numbers too, as Excel does
 * ("Greater Than 10" keeps 2 to 9 as well as 11 and up).
 */
function labelTest(filter: Extract<PivotFilter, { kind: "label" }>): (v: CellValue) => boolean {
  const text = (v: CellValue) => labelText(v).toLowerCase()
  const a = filter.value
  const b = filter.value2 ?? ""
  const cmp = (v: CellValue, t: string) => text(v).localeCompare(t.toLowerCase(), undefined, { sensitivity: "base" })
  const like = (pattern: string) => {
    const re = wildcardRegExp(pattern.toLowerCase())
    return (v: CellValue) => re.test(text(v))
  }
  const literal = (t: string) => t.replace(/[~*?]/g, (c) => `~${c}`)
  switch (filter.op) {
    case "equal":
      return like(a)
    case "notEqual":
      return (v) => !like(a)(v)
    case "beginsWith":
      return like(`${literal(a)}*`)
    case "notBeginsWith":
      return (v) => !like(`${literal(a)}*`)(v)
    case "endsWith":
      return like(`*${literal(a)}`)
    case "notEndsWith":
      return (v) => !like(`*${literal(a)}`)(v)
    case "contains":
      return like(`*${literal(a)}*`)
    case "notContains":
      return (v) => !like(`*${literal(a)}*`)(v)
    case "greaterThan":
      return (v) => cmp(v, a) > 0
    case "greaterThanOrEqual":
      return (v) => cmp(v, a) >= 0
    case "lessThan":
      return (v) => cmp(v, a) < 0
    case "lessThanOrEqual":
      return (v) => cmp(v, a) <= 0
    case "between":
      return (v) => cmp(v, a) >= 0 && cmp(v, b) <= 0
    case "notBetween":
      return (v) => cmp(v, a) < 0 || cmp(v, b) > 0
  }
}

function valueTest(filter: Extract<PivotFilter, { kind: "value" }>): (v: number) => boolean {
  const a = filter.value
  const b = filter.value2 ?? a
  const [lo, hi] = a <= b ? [a, b] : [b, a]
  switch (filter.op) {
    case "equal":
      return (v) => v === a
    case "notEqual":
      return (v) => v !== a
    case "greaterThan":
      return (v) => v > a
    case "greaterThanOrEqual":
      return (v) => v >= a
    case "lessThan":
      return (v) => v < a
    case "lessThanOrEqual":
      return (v) => v <= a
    case "between":
      return (v) => v >= lo && v <= hi
    case "notBetween":
      return (v) => v < lo || v > hi
  }
}

/** The records each axis lays its lines out from, and those the values count (in both). */
type Filtered = { rows: number[]; cols: number[]; both: number[] }

/**
 * The records left once the table's filters have had their say. A label filter drops the items whose label fails,
 * everywhere (like unticked items). A value filter or Top 10 judges each item by its total (of the filter's value
 * field, over the records still in) under the items of the fields before it on its axis, and only narrows that axis:
 * Excel keeps the other axis' items even when a column filter leaves them no values.
 */
function applyFilters(
  table: PivotTable,
  cache: PivotCache,
  items: Map<number, FieldItems>,
  records: number[],
): Filtered {
  const filters = table.filters ?? []
  let kept = records
  for (const filter of filters) {
    if (filter.kind !== "label") continue
    const fi = items.get(filter.field)
    if (!fi) continue
    const test = labelTest(filter)
    const pass = new Map<number, boolean>()
    kept = kept.filter((r) => {
      const pos = fi.ofRecord[r]
      let ok = pass.get(pos)
      if (ok === undefined) pass.set(pos, (ok = test(itemLabel(fi, pos))))
      return ok
    })
  }
  const axisKept = (axis: number[]) => {
    let out = kept
    for (const filter of filters)
      if ((filter.kind === "value" || filter.kind === "top") && axis.includes(filter.field))
        out = measureFilter(table, cache, items, out, axis, filter)
    return out
  }
  const rows = axisKept(table.rows)
  const cols = axisKept(table.cols)
  if (rows === kept && cols === kept) return { rows, cols, both: kept }
  const inCols = new Set(cols)
  return { rows, cols, both: rows.filter((r) => inCols.has(r)) }
}

/** The records whose item of `filter.field` passes a value filter or Top 10, judged within its parent items. */
function measureFilter(
  table: PivotTable,
  cache: PivotCache,
  items: Map<number, FieldItems>,
  records: number[],
  axis: number[],
  filter: Extract<PivotFilter, { kind: "value" | "top" }>,
): number[] {
  const fi = items.get(filter.field)
  const df = table.values[filter.measure]
  if (!fi || !df) return records
  const real = axis.filter((f) => f !== VALUES_FIELD)
  const outer = real.slice(0, real.indexOf(filter.field))
  const parentOf = (r: number) => outer.map((f) => items.get(f)!.ofRecord[r]).join(",")
  // Each item's total under its parent items.
  const inputs = calcFieldInputs(cache, df.field)
  const groups = new Map<string, Map<number, Map<number, Stats>>>()
  for (const r of records) {
    const key = parentOf(r)
    let g = groups.get(key)
    if (!g) groups.set(key, (g = new Map()))
    const pos = fi.ofRecord[r]
    let st = g.get(pos)
    if (!st) g.set(pos, (st = new Map(inputs.map((f) => [f, newStats()]))))
    for (const f of inputs) addValue(st.get(f)!, cache.records[r][f] ?? null)
  }
  const keep = new Set<string>()
  for (const [key, g] of groups) {
    const totals: { pos: number; v: number }[] = []
    for (const [pos, st] of g) {
      const v = cache.fields[df.field]?.formula
        ? calcFieldValue(cache, df.field, (g) => st.get(g)?.sum ?? 0)
        : summarize(st.get(df.field)!, df.fn)
      if (typeof v === "number") totals.push({ pos, v })
    }
    if (filter.kind === "value") {
      const test = valueTest(filter)
      for (const t of totals) if (test(t.v)) keep.add(`${key}|${t.pos}`)
      continue
    }
    // Top 10: biggest first (smallest first for Bottom).
    const ranked = totals.sort((x, y) => (filter.top ? y.v - x.v : x.v - y.v))
    if (filter.by === "count") {
      const n = Math.max(0, Math.floor(filter.value))
      ranked.slice(0, n).forEach((t) => keep.add(`${key}|${t.pos}`))
    } else {
      // Percent and Sum: items in turn until they reach the share of the whole or the sum asked for.
      const whole = ranked.reduce((sum, t) => sum + t.v, 0)
      const target = filter.by === "percent" ? (whole * filter.value) / 100 : filter.value
      let sum = 0
      for (const t of ranked) {
        if (sum >= target) break
        keep.add(`${key}|${t.pos}`)
        sum += t.v
      }
    }
  }
  return records.filter((r) => keep.has(`${parentOf(r)}|${fi.ofRecord[r]}`))
}

/**
 * The summary figures of every group the layout needs, gathered in one pass over the records: for each record, every
 * prefix of its row items crossed with every prefix of its column items.
 */
function makeStats(table: PivotTable, cache: PivotCache, items: Map<number, FieldItems>, records: number[]) {
  const realRows = table.rows.filter((f) => f !== VALUES_FIELD)
  const realCols = table.cols.filter((f) => f !== VALUES_FIELD)
  // A calculated field reads the fields its formula names.
  const sources = [...new Set(table.values.flatMap((v) => calcFieldInputs(cache, v.field)))]
  const groups = new Map<string, Map<number, Stats>>()
  for (const r of records) {
    const rowPos = realRows.map((f) => items.get(f)!.ofRecord[r])
    const colPos = realCols.map((f) => items.get(f)!.ofRecord[r])
    for (let a = 0; a <= rowPos.length; a++) {
      const rk = rowPos.slice(0, a).join(",")
      for (let b = 0; b <= colPos.length; b++) {
        const key = `${rk}|${colPos.slice(0, b).join(",")}`
        let g = groups.get(key)
        if (!g) groups.set(key, (g = new Map()))
        for (const f of sources) {
          let s = g.get(f)
          if (!s) g.set(f, (s = newStats()))
          addValue(s, cache.records[r][f] ?? null)
        }
      }
    }
  }
  /** A line's items on the real fields (the Values level left out); a grand total line has none. */
  const realPath = (line: Line, fields: number[]) =>
    line.t === "grand" ? [] : line.path.filter((_, j) => fields[j] !== VALUES_FIELD)
  /** The figures for row items `rows` × column items `cols` (positions per real field). */
  const at = (rows: readonly number[], cols: readonly number[], field: number): Stats =>
    groups.get(`${rows.join(",")}|${cols.join(",")}`)?.get(field) ?? newStats()
  const forLines = (row: Line, rowFields: number[], col: Line, colFields: number[], field: number): Stats =>
    at(realPath(row, rowFields), realPath(col, colFields), field)
  return Object.assign(forLines, { at, realPath })
}

// ── Calculated items ──

type Measure = (field: number, fn: PivotFn, rows: readonly number[], cols: readonly number[]) => CellValue | undefined

/** A total of item values put together as `fn` does (Sum and Count add, Max and Min pick); errors carry through. */
function combine(fn: PivotFn, values: (CellValue | undefined)[]): CellValue | undefined {
  let out: number | undefined
  for (const v of values) {
    if (v === undefined || v === null) continue
    if (typeof v === "object") return v
    if (typeof v !== "number") return errorValue("#VALUE!")
    if (out === undefined) out = v
    else if (fn === "max") out = Math.max(out, v)
    else if (fn === "min") out = Math.min(out, v)
    else if (fn === "product") out *= v
    else out += v
  }
  return out
}

/**
 * `plain` (a value field's figure from the records) with the table's calculated items worked out: a cell on a
 * calculated item takes its formula over the other items' figures in that cell (the item last in solve order, where
 * two meet); a total with calculated items inside it puts its items' figures together. `plain` itself when the
 * table shows no calculated item.
 */
function withCalcItems(
  cache: PivotCache,
  items: Map<number, FieldItems>,
  rowFields: number[],
  colFields: number[],
  records: number[],
  plain: Measure,
): Measure {
  const calcRows = rowFields.map((f) => items.get(f)!.calc)
  const calcCols = colFields.map((f) => items.get(f)!.calc)
  if (![...calcRows, ...calcCols].some((m) => m.size)) return plain
  const keyOf = (path: readonly number[]) => path.map((p) => `${p},`).join("")
  /** The items with data under each path of items on an axis. */
  const childrenOf = (fields: number[]) => {
    const map = new Map<string, Set<number>>()
    for (const r of records) {
      let key = ""
      for (const f of fields) {
        const pos = items.get(f)!.ofRecord[r]
        let set = map.get(key)
        if (!set) map.set(key, (set = new Set()))
        set.add(pos)
        key += `${pos},`
      }
    }
    return map
  }
  const rowKids = childrenOf(rowFields)
  const colKids = childrenOf(colFields)
  /** The items with data under `path` (under a calculated item: those under any of its siblings). */
  const realKids = (map: Map<string, Set<number>>, path: readonly number[], fields: number[]): Set<number> => {
    const j = path.findIndex((pos, k) => items.get(fields[k])!.calc.has(pos))
    if (j < 0) return map.get(keyOf(path)) ?? new Set()
    const out = new Set<number>()
    for (const sibling of map.get(keyOf(path.slice(0, j))) ?? []) {
      const next = [...path]
      next[j] = sibling
      for (const pos of realKids(map, next, fields)) out.add(pos)
    }
    return out
  }
  /** The items a total over `path` puts together: those with data, and the next field's shown calculated items. */
  const kids = (map: Map<string, Set<number>>, path: readonly number[], fields: number[]) => {
    const fi = items.get(fields[path.length])!
    const out = [...realKids(map, path, fields)]
    for (const pos of fi.calc.keys()) if (!out.includes(pos) && fi.order.includes(pos)) out.push(pos)
    return out
  }
  const memo = new Map<string, CellValue | undefined>()
  const value: Measure = (field, fn, rows, cols) => {
    const key = `${field}|${fn}|${keyOf(rows)}|${keyOf(cols)}`
    if (memo.has(key)) return memo.get(key)
    // A formula that comes back to itself.
    memo.set(key, errorValue("#REF!"))
    const v = compute(field, fn, rows, cols)
    memo.set(key, v)
    return v
  }
  const calcIn = (path: readonly number[], calc: Map<number, number>[]) => path.some((pos, j) => calc[j].has(pos))
  const calcBelow = (path: readonly number[], calc: Map<number, number>[]) =>
    calc.slice(path.length).some((m) => m.size > 0)
  const compute: Measure = (field, fn, rows, cols) => {
    const onCalc = calcIn(rows, calcRows) || calcIn(cols, calcCols)
    if (!onCalc && !calcBelow(rows, calcRows) && !calcBelow(cols, calcCols)) return plain(field, fn, rows, cols)
    // Excel works formulas out in the innermost cells only: a total puts its cells together, whatever it's on.
    if (CALC_ITEM_FNS.includes(fn)) {
      if (rows.length < rowFields.length)
        return combine(
          fn,
          kids(rowKids, rows, rowFields).map((pos) => value(field, fn, [...rows, pos], cols)),
        )
      if (cols.length < colFields.length)
        return combine(
          fn,
          kids(colKids, cols, colFields).map((pos) => value(field, fn, rows, [...cols, pos])),
        )
    }
    if (!onCalc) return plain(field, fn, rows, cols)
    // The calculated item last in solve order among the cell's items works it out.
    let best: { row: boolean; j: number; k: number } | null = null
    const look = (path: readonly number[], calc: Map<number, number>[], row: boolean) =>
      path.forEach((pos, j) => {
        const k = calc[j].get(pos)
        if (k !== undefined && (!best || k > best.k)) best = { row, j, k }
      })
    look(rows, calcRows, true)
    look(cols, calcCols, false)
    const { row, j, k } = best as unknown as { row: boolean; j: number; k: number }
    const fi = items.get((row ? rowFields : colFields)[j])!
    const path = row ? rows : cols
    const at = (pos: number) => {
      const next = [...path]
      next[j] = pos
      return row ? value(field, fn, next, cols) : value(field, fn, rows, next)
    }
    return calcItemValue(cache, k, (ref) => {
      if (ref.kind === "item") {
        const pos = fi.xs.indexOf(ref.x)
        return pos < 0 ? undefined : at(pos)
      }
      const i = ref.kind === "at" ? ref.n - 1 : fi.order.indexOf(path[j]) + ref.n
      return i >= 0 && i < fi.order.length ? at(fi.order[i]) : errorValue("#REF!")
    })
  }
  return value
}
