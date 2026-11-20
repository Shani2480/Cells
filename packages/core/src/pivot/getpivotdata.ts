// GETPIVOTDATA: the value cell of a PivotTable that a value field and field/item pairs name. Excel 365's rules,
// checked formula by formula against Excel on io's pivot fixtures (io/test/getpivotdata.test.ts): names match
// ignoring case (a value field by its own name or its source field's), pairs in any order; the items must pick out
// one line the PivotTable shows on each axis (an outer field may be left out when only one of its items has the
// rest); a report filter field counts only when it shows that one item; a line the PivotTable doesn't show
// (hidden, filtered, collapsed, no subtotal) or more than one is #REF!.
import { columnName, type Range } from "../address"
import { serialToDate } from "../formula/dates"
import { pivotAxes, pivotFieldName, pivotLayout, type PivotAxisLine, type PivotLayout } from "./engine"
import { VALUES_FIELD, type PivotCache, type PivotTable, type PivotValue } from "./model"

export type PivotLooked = {
  layout: PivotLayout
  axes: ReturnType<typeof pivotAxes>
  cells: Map<string, { row: number; col: number }>
}

/** Layouts per PivotTable object (tables are replaced, never changed, so the object is the key). */
const looked = new WeakMap<PivotTable, { cache: PivotCache; at: PivotLooked }>()

/** The layout, axes and value cells of a PivotTable (kept per table object). */
export function pivotLook(table: PivotTable, cache: PivotCache): PivotLooked {
  const hit = looked.get(table)
  if (hit && hit.cache === cache) return hit.at
  const layout = pivotLayout(table, cache)
  const cells = new Map<string, { row: number; col: number }>()
  for (const c of layout.cells) if (c.lines) cells.set(`${c.lines[0]},${c.lines[1]}`, { row: c.row, col: c.col })
  const at = { layout, axes: pivotAxes(table, cache), cells }
  looked.set(table, { cache, at })
  return at
}

/** The PivotTable's cells with its report filters above them: where a GETPIVOTDATA reference may point. */
export function pivotTableRange(table: PivotTable, cache: PivotCache): Range {
  const { range, pages } = pivotLook(table, cache).layout
  return pages ? { ...range, r0: pages.r0 } : range
}

/**
 * The cell holding value field `dataField` for the field/item `pairs`, or null (#REF!). Items are text, numbers or
 * logicals as the formula gives them.
 */
export function pivotDataCell(
  table: PivotTable,
  cache: PivotCache,
  dataField: string,
  pairs: readonly [string, string | number | boolean][],
): { row: number; col: number } | null {
  const { axes, cells } = pivotLook(table, cache)
  const lower = dataField.toLowerCase()
  let data = table.values.findIndex((v) => v.name.toLowerCase() === lower)
  if (data < 0)
    data = table.values.findIndex(
      (v) =>
        pivotFieldName(table, cache, v.field).toLowerCase() === lower ||
        cache.fields[v.field]?.name.toLowerCase() === lower,
    )
  if (data < 0) return null

  const fieldOf = (name: string) => {
    const want = name.toLowerCase()
    const known = [...axes.rowFields, ...axes.colFields, ...table.pages.map((p) => p.field)]
    return known.find((f) => pivotFieldName(table, cache, f).toLowerCase() === want)
  }
  const rows = new Map<number, number>()
  const cols = new Map<number, number>()
  for (const [name, item] of pairs) {
    const f = fieldOf(name)
    if (f === undefined) return null
    const pos = axes.itemOf(f, item)
    if (pos < 0) return null
    const page = table.pages.find((p) => p.field === f)
    if (page) {
      if (page.item !== pos) return null
      continue
    }
    const target = axes.rowFields.includes(f) ? rows : cols
    if (target.has(f)) return null
    target.set(f, pos)
  }
  const dataOnRows = table.rows.includes(VALUES_FIELD)
  const dataOnCols = table.cols.includes(VALUES_FIELD)
  /**
   * The one line of an axis the items asked for pick out: a line as deep as the innermost field asked for, with
   * those items (fields left out may be anything, but only one line may match); with none asked, the grand total.
   * -1 when no line or more than one matches.
   */
  const find = (lines: PivotAxisLine[], fields: number[], asked: Map<number, number>, dataHere: boolean) => {
    const levels = [...asked].map(([f, pos]) => [fields.indexOf(f), pos] as const)
    const depth = levels.length ? Math.max(...levels.map(([j]) => j)) + 1 : 0
    let found = -1
    lines.forEach((l, i) => {
      if (!l.values || (dataHere && l.data !== data) || l.items.length !== depth) return
      if (depth ? l.t !== "data" && l.t !== "default" : fields.length ? l.t !== "grand" : l.t !== "data") return
      if (!levels.every(([j, pos]) => l.items[j] === pos)) return
      found = found === -1 ? i : -2
    })
    return found
  }
  const ri = find(axes.rowLines, axes.rowFields, rows, dataOnRows)
  const ci = find(axes.colLines, axes.colFields, cols, dataOnCols)
  if (ri < 0 || ci < 0) return null
  return cells.get(`${ri},${ci}`) ?? null
}

/** A value of a GETPIVOTDATA formula: text quoted, dates as DATE(y,m,d), numbers and logicals as they are. */
function argText(v: PivotValue, date: boolean): string {
  if (typeof v === "number") {
    if (!date) return String(v)
    const d = serialToDate(v)
    return `DATE(${d.year},${d.month},${d.day})`
  }
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE"
  if (v === null || v === "") return '"(blank)"'
  const text = typeof v === "object" ? v.error : v
  return `"${text.replace(/"/g, '""')}"`
}

/**
 * Generate GetPivotData: the GETPIVOTDATA call Excel writes when a formula points at a value cell of a PivotTable
 * (without the "="): the value field, the table's top-left cell, and each row and column item of the cell. Null
 * for a cell that isn't a value.
 */
export function getPivotDataFormula(table: PivotTable, cache: PivotCache, row: number, col: number): string | null {
  const { layout, axes } = pivotLook(table, cache)
  const cell = layout.cells.find((c) => c.row === row && c.col === col)
  if (!cell?.lines) return null
  const [ri, ci] = cell.lines
  const rowLine = axes.rowLines[ri]
  const colLine = axes.colLines[ci]
  const onRows = table.rows.includes(VALUES_FIELD)
  const onCols = table.cols.includes(VALUES_FIELD)
  const df = table.values[onRows ? rowLine.data : onCols ? colLine.data : 0]
  if (!df) return null
  const at = `$${columnName(layout.range.c0)}$${layout.range.r0 + 1}`
  const parts = [argText(df.name, false), at]
  const asked: [string, string | number | boolean][] = []
  const pairs = (fields: number[], line: PivotAxisLine) =>
    line.items.forEach((pos, j) => {
      const f = fields[j]
      const name = pivotFieldName(table, cache, f)
      const item = axes.itemValue(f, pos)
      parts.push(argText(name, false), argText(item, Boolean(cache.fields[f]?.dates)))
      asked.push([name, item === null || item === "" ? "(blank)" : typeof item === "object" ? item.error : item])
    })
  pairs(axes.rowFields, rowLine)
  pairs(axes.colFields, colLine)
  // Cells GETPIVOTDATA can't name (a custom subtotal's, say) keep a plain reference, as the call would be #REF!.
  const back = pivotDataCell(table, cache, df.name, asked)
  if (back?.row !== row || back.col !== col) return null
  return `GETPIVOTDATA(${parts.join(",")})`
}
