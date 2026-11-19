import type { Range } from "../address"
import { errorValue, type CellError, type CellValue } from "../values"
import {
  PIVOT_FN_LABEL,
  VALUES_FIELD,
  type PivotCache,
  type PivotFn,
  type PivotSubtotal,
  type PivotTable,
  type PivotValue,
} from "./model"

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

const keyOf = (v: PivotValue): string =>
  v === null || v === ""
    ? "blank"
    : typeof v === "number"
      ? `n:${v}`
      : typeof v === "string"
        ? `s:${v.toLowerCase()}`
        : typeof v === "boolean"
          ? `b:${v}`
          : `e:${v.error}`

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
}

function fieldItems(table: PivotTable, cache: PivotCache, f: number): FieldItems {
  const field = table.fields[f]
  const cacheField = cache.fields[f]
  // The cache's shared items, or the column's distinct values when it keeps none.
  let shared = cacheField.items
  const sharedKey = new Map<string, number>()
  if (shared) shared.forEach((v, x) => sharedKey.has(keyOf(v)) || sharedKey.set(keyOf(v), x))
  else {
    shared = []
    for (const rec of cache.records) {
      const k = keyOf(rec[f] ?? null)
      if (!sharedKey.has(k)) {
        sharedKey.set(k, shared.length)
        shared.push(rec[f] ?? null)
      }
    }
  }
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
  const ofRecord = new Int32Array(cache.records.length)
  cache.records.forEach((rec, i) => {
    ofRecord[i] = posOfShared.get(sharedKey.get(keyOf(rec[f] ?? null)) ?? -1) ?? -1
  })
  let order = xs.map((_, pos) => pos).filter((pos) => !hidden.has(pos))
  if (field?.sort === "ascending" || field?.sort === "descending") {
    const sign = field.sort === "ascending" ? 1 : -1
    order = order.sort((a, b) => {
      const va = names[a] ?? values[a]
      const vb = names[b] ?? values[b]
      // Blanks stay last either way.
      const blanks = Number(rank(va) === 4) - Number(rank(vb) === 4)
      return blanks || sign * comparePivotValues(va, vb)
    })
  }
  return { values, xs, order, hidden, collapsed, names, ofRecord }
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

/** A group of records under one path of items. */
type Node = { children: Map<number, Node> }

function axisLines(
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
  const root: Node = { children: new Map() }
  for (const r of records) {
    let node = root
    for (const f of real) {
      const pos = items.get(f)!.ofRecord[r]
      let next = node.children.get(pos)
      if (!next) node.children.set(pos, (next = { children: new Map() }))
      node = next
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
      const shown = field.showAll ? fi.order : fi.order.filter((pos) => node.children.has(pos))
      children = shown.map((pos) => ({ pos, node: node.children.get(pos) ?? { children: new Map() } }))
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

export function pivotLayout(table: PivotTable, cache: PivotCache): PivotLayout {
  const rowFields = table.rows
  const colFields = table.cols
  const items = new Map<number, FieldItems>()
  for (const f of [...rowFields, ...colFields, ...table.pages.map((p) => p.field)])
    if (f !== VALUES_FIELD && !items.has(f) && cache.fields[f]) items.set(f, fieldItems(table, cache, f))

  // The records that count: report filters' picks, and no hidden item on a row or column field.
  const records: number[] = []
  next: for (let r = 0; r < cache.records.length; r++) {
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

  const rowLines = axisLines(table, rowFields, items, records, true)
  const colLines = axisLines(table, colFields, items, records, false)

  // Label columns: a compact field shares its column with the field after it.
  const compactOf = (f: number) =>
    f === VALUES_FIELD ? table.compactData : table.fields[f].compact && table.fields[f].outline
  const labelCol: number[] = []
  rowFields.forEach((_f, j) => labelCol.push(j === 0 ? 0 : labelCol[j - 1] + (compactOf(rowFields[j - 1]) ? 0 : 1)))
  const labelCols = rowFields.length ? labelCol[labelCol.length - 1] + 1 : 0
  const n = colFields.length
  const headerRows = n ? n + 1 : 1
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
    if (first && table.showDrill !== false && fields.slice(level + 1).some((g) => g !== VALUES_FIELD))
      extra.button = fi.collapsed.has(pos) ? "expand" : "collapse"
    return extra
  }

  // ── The header rows ──
  const dataInRows = rowFields.includes(VALUES_FIELD)
  const dataInCols = colFields.includes(VALUES_FIELD)
  if (n) {
    // With no row fields the column header takes the corner's cell.
    if (!labelCols) {
      /* no corner */
    } else if (table.values.length === 1 && !dataInRows && !dataInCols) put(0, 0, "corner", dataName(0))
    else put(0, 0, "corner")
    if (colFields.length === 1 && dataInCols) put(0, labelCols, "colHeader", table.captions.data)
    else if (compactOf(colFields[0])) put(0, labelCols, "colHeader", table.captions.colHeader ?? "Column Labels")
    else colFields.forEach((f, j) => put(0, labelCols + j, "colHeader", pivotFieldName(table, cache, f)))
  }
  const headerRow = headerRows - 1
  const rowHeaderCaption = table.captions.rowHeader ?? "Row Labels"
  for (let c = 0; c < labelCols; c++) {
    const j = labelCol.indexOf(c)
    const f = rowFields[j]
    const shared = labelCol.filter((x) => x === c).length > 1
    put(headerRow, c, "rowHeader", shared || compactOf(f) ? rowHeaderCaption : pivotFieldName(table, cache, f))
  }
  if (!n && table.values.length === 1) put(0, labelCols, "dataHeader", dataName(0))

  // Column labels: each level's item over the first column of its group.
  const colLevelRow = (j: number) => 1 + j
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
  const statsFor = makeStats(table, cache, items, records)
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
      let value = summarize(statsFor(line, rowFields, colLine, colFields, df.field), fn)
      if (value !== undefined && typeof value === "object" && table.showError && table.captions.error !== undefined)
        value = table.captions.error
      if (value === undefined && table.showMissing && table.captions.missing) value = table.captions.missing
      put(row, col, role, value, df.numFmt ? { numFmt: df.numFmt } : undefined)
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
  return {
    range: { r0: top, c0: left, r1: top + height - 1, c1: left + Math.max(width, 1) - 1 },
    pages: pageCount ? { r0: top - pageCount - 1, c0: left, r1: top - 2, c1: left + 1 } : null,
    firstHeaderRow: 1,
    firstDataRow: headerRows,
    firstDataCol: labelCols,
    rowLabelCols: labelCol,
    cells,
    rowItems: axisItems(rowLines, dataInRows),
    colItems: axisItems(colLines, dataInCols),
  }
}

/**
 * The summary figures of every group the layout needs, gathered in one pass over the records: for each record, every
 * prefix of its row items crossed with every prefix of its column items.
 */
function makeStats(table: PivotTable, cache: PivotCache, items: Map<number, FieldItems>, records: number[]) {
  const realRows = table.rows.filter((f) => f !== VALUES_FIELD)
  const realCols = table.cols.filter((f) => f !== VALUES_FIELD)
  const sources = [...new Set(table.values.map((v) => v.field))]
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
  return (row: Line, rowFields: number[], col: Line, colFields: number[], field: number): Stats => {
    const key = `${realPath(row, rowFields).join(",")}|${realPath(col, colFields).join(",")}`
    return groups.get(key)?.get(field) ?? newStats()
  }
}
