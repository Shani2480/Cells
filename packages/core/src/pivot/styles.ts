import type { Range } from "../address"
import { OFFICE_THEME, sameLook, type Theme } from "../theme"
import { rethemeStyleDef, type TableCellLook, type TableStyleElement } from "../table/styles"
import { BUILT_IN_PIVOT_STYLES } from "./built-in-styles"
import { pivotLayout, type PivotCell, type PivotLayout } from "./engine"
import type { PivotCache, PivotTable } from "./model"

/**
 * PivotTable styles (PivotTable Design › PivotTable Styles): Excel's 84 built-in styles, and the look each gives a
 * cell of a PivotTable, worked out from where the cell sits in the layout.
 */

export type PivotStyleElementName =
  | "wholeTable"
  | "headerRow"
  /** The Grand Total row. */
  | "totalRow"
  | "firstColumn"
  /** The Grand Total column. */
  | "lastColumn"
  | "firstRowStripe"
  | "secondRowStripe"
  | "firstColumnStripe"
  | "secondColumnStripe"
  | "firstHeaderCell"
  | "subtotalColumn1"
  | "subtotalColumn2"
  | "subtotalColumn3"
  | "subtotalRow1"
  | "subtotalRow2"
  | "subtotalRow3"
  | "blankRow"
  | "columnSubheading1"
  | "columnSubheading2"
  | "columnSubheading3"
  | "rowSubheading1"
  | "rowSubheading2"
  | "rowSubheading3"
  | "pageFieldLabels"
  | "pageFieldValues"

export type PivotStyleDef = Partial<Record<PivotStyleElementName, TableStyleElement>>

/** A PivotTable style's look on a cell, with the indent of compact-form labels. */
export type PivotCellLook = TableCellLook & { indent?: number }

/** The built-in styles in Excel's gallery order, by group. */
export const PIVOT_STYLE_GROUPS: { label: string; names: string[] }[] = ["Light", "Medium", "Dark"].map((label) => ({
  label,
  names: Array.from({ length: 28 }, (_, i) => `PivotStyle${label}${i + 1}`),
}))

const themed = new WeakMap<Theme, Map<string, PivotStyleDef>>()

/** The built-in style for a name ("PivotStyleLight16"), or null; with `theme`, in that theme's colours. */
export function pivotStyleDef(name: string | null, theme?: Theme): PivotStyleDef | null {
  if (!name) return null
  const def = BUILT_IN_PIVOT_STYLES[name.replace(/^PivotStyle/i, "")] ?? null
  if (!def || !theme || sameLook(theme, OFFICE_THEME)) return def
  let byName = themed.get(theme)
  if (!byName) themed.set(theme, (byName = new Map()))
  let out = byName.get(name)
  if (!out) byName.set(name, (out = rethemeStyleDef(def, theme)))
  return out
}

const layouts = new WeakMap<
  PivotTable,
  { cache: PivotCache; layout: PivotLayout; lines: Lines; cells: Map<string, PivotCell> }
>()

/** What each row and column of the layout is, for the style. */
type Lines = {
  /** Per body row: its line kind and level. */
  rows: { t: string; level: number; outer: boolean }[]
  cols: { t: string; level: number; outer: boolean }[]
  indent: Map<string, number>
  /**
   * Row subheadings in their own label column (tabular and outline-without-compact levels): "row,col" of each cell
   * to the rows of its item's group, which the subheading covers.
   */
  groups: Map<string, { r0: number; r1: number; level: number }>
}

/** The engine's layout of a PivotTable, kept while the table and its cache are the same objects. */
export function pivotLayoutOf(table: PivotTable, cache: PivotCache): PivotLayout {
  return memo(table, cache).layout
}

/** The engine's cell of a PivotTable at a sheet cell (what it shows and stands for), if the table has one there. */
export function pivotCellAt(table: PivotTable, cache: PivotCache, row: number, col: number): PivotCell | undefined {
  return memo(table, cache).cells.get(`${row},${col}`)
}

function memo(table: PivotTable, cache: PivotCache) {
  const hit = layouts.get(table)
  if (hit && hit.cache === cache) return hit
  const layout = pivotLayout(table, cache)
  const rowDepth = table.rows.length
  const colDepth = table.cols.length
  const indent = new Map<string, number>()
  for (const c of layout.cells) if (c.indent) indent.set(`${c.row},${c.col}`, c.indent)
  const level = (i: { t: string; r: number; x: number[] }) =>
    (i.t === "data" || i.t === "blank" ? i.r : 0) + i.x.length - 1
  const lines: Lines = {
    rows: layout.rowItems.map((i) => ({ t: i.t, level: level(i), outer: i.t === "data" && level(i) < rowDepth - 1 })),
    cols: layout.colItems.map((i) => ({ t: i.t, level: level(i), outer: i.t === "data" && level(i) < colDepth - 1 })),
    indent,
    groups: new Map(),
  }
  // A level whose next level has a column of its own shows its labels beside the inner ones, on their rows: its
  // subheading covers its label column down the group's rows.
  const lc = layout.rowLabelCols
  const top = layout.range.r0 + layout.firstDataRow
  const items = layout.rowItems
  for (let j = 0; j < rowDepth - 1; j++) {
    if (lc[j + 1] === lc[j]) continue
    const col = layout.range.c0 + lc[j]
    for (let k = 0; k < items.length; k++) {
      const i = items[k]
      if (i.t !== "data" || i.r > j || level(i) <= j) continue
      let end = k
      while (end + 1 < items.length && items[end + 1].t === "data" && items[end + 1].r > j) end++
      for (let m = k; m <= end; m++) lines.groups.set(`${top + m},${col}`, { r0: top + k, r1: top + end, level: j })
      k = end
    }
  }
  const out = { cache, layout, lines, cells: new Map(layout.cells.map((c) => [`${c.row},${c.col}`, c])) }
  layouts.set(table, out)
  return out
}

const sub = (prefix: string, level: number) => `${prefix}${Math.min(level, 2) + 1}` as PivotStyleElementName

/**
 * The look a PivotTable's style gives one of its cells (null outside it, or without a style). Elements apply in
 * Excel's order, each over the ones before: whole table, report filters, column and row stripes, the label column,
 * the header rows, the first header cell, subtotal columns, blank rows, subtotal rows, column and row subheadings,
 * the Grand Total column, the Grand Total row. An element's outer borders go on the edges of the cells it covers.
 */
export function pivotCellLook(
  table: PivotTable,
  cache: PivotCache,
  row: number,
  col: number,
  def = pivotStyleDef(table.style.name),
): PivotCellLook | null {
  const { layout, lines } = memo(table, cache)
  const L = layout.range
  const out: PivotCellLook = {}
  const ind = lines.indent.get(`${row},${col}`)
  if (ind) out.indent = ind
  if (!def) return ind ? out : null
  const apply = (e: TableStyleElement | undefined, box: Range) => {
    if (!e) return
    if (e.fill) out.fill = e.fill
    if (e.bold) out.bold = true
    if (e.color) out.fontColor = e.color
    const b = e.border
    if (!b) return
    const top = row === box.r0 ? b.top : b.horizontal
    const bottom = row === box.r1 ? b.bottom : b.horizontal
    const left = col === box.c0 ? b.left : b.vertical
    const right = col === box.c1 ? b.right : b.vertical
    if (top) out.borderTop = top
    if (bottom) out.borderBottom = bottom
    if (left) out.borderLeft = left
    if (right) out.borderRight = right
  }
  // The report filters, above the table.
  if (layout.pages && row >= layout.pages.r0 && row <= layout.pages.r1) {
    if (col === layout.pages.c0) apply(def.pageFieldLabels, { ...layout.pages, c1: col })
    else if (col === layout.pages.c0 + 1) apply(def.pageFieldValues, { ...layout.pages, c0: col })
    return out
  }
  if (row < L.r0 || row > L.r1 || col < L.c0 || col > L.c1) return ind ? out : null
  const body0 = L.r0 + layout.firstDataRow
  const data0 = L.c0 + layout.firstDataCol
  const inBody = row >= body0
  const header: Range = { r0: L.r0, r1: body0 - 1, c0: L.c0, c1: L.c1 }
  const body: Range = { r0: body0, r1: L.r1, c0: L.c0, c1: L.c1 }
  const line = inBody ? lines.rows[row - body0] : undefined
  const colLine = col >= data0 ? lines.cols[col - data0] : undefined
  const s = table.style

  apply(def.wholeTable, L)
  if (s.colStripes && inBody && col >= data0) {
    const k = col - data0
    const a = def.firstColumnStripe?.stripe ?? 1
    const b = def.secondColumnStripe?.stripe ?? 1
    const at = k % (a + b)
    const c0 = col - at
    if (at < a) apply(def.firstColumnStripe, { ...body, c0, c1: c0 + a - 1 })
    else apply(def.secondColumnStripe, { ...body, c0: c0 + a, c1: c0 + a + b - 1 })
  }
  if (s.rowStripes && inBody) {
    const k = row - body0
    const a = def.firstRowStripe?.stripe ?? 1
    const b = def.secondRowStripe?.stripe ?? 1
    const at = k % (a + b)
    const r0 = row - at
    if (at < a) apply(def.firstRowStripe, { ...body, r0, r1: r0 + a - 1 })
    else apply(def.secondRowStripe, { ...body, r0: r0 + a, r1: r0 + a + b - 1 })
  }
  if (s.rowHeaders && inBody && col < data0) apply(def.firstColumn, { ...body, c1: data0 - 1 })
  if (s.colHeaders && !inBody) apply(def.headerRow, header)
  if (s.rowHeaders && row === L.r0 && col === L.c0) apply(def.firstHeaderCell, { r0: row, r1: row, c0: col, c1: col })
  if (colLine && colLine.t !== "data" && colLine.t !== "grand")
    apply(def[sub("subtotalColumn", colLine.level)], { r0: L.r0 + 1 + colLine.level, r1: L.r1, c0: col, c1: col })
  if (line?.t === "blank") apply(def.blankRow, { r0: row, r1: row, c0: L.c0, c1: L.c1 })
  // Only real subtotal lines: an outer item's own line (totals at the top or not) is a row subheading.
  if (line && line.t !== "data" && line.t !== "grand" && line.t !== "blank")
    apply(def[sub("subtotalRow", line.level)], { r0: row, r1: row, c0: L.c0, c1: L.c1 })
  if (!inBody && col >= data0 && table.cols.length > 1) {
    const level = row - L.r0 - 1
    if (level >= 0 && level < table.cols.length - 1)
      apply(def[sub("columnSubheading", level)], { r0: row, r1: row, c0: data0, c1: L.c1 })
  }
  if (line?.outer) apply(def[sub("rowSubheading", line.level)], { r0: row, r1: row, c0: L.c0, c1: L.c1 })
  const group = lines.groups.get(`${row},${col}`)
  if (group) apply(def[sub("rowSubheading", group.level)], { ...group, c0: col, c1: col })
  if (s.lastColumn && colLine?.t === "grand") apply(def.lastColumn, { r0: L.r0 + 1, r1: L.r1, c0: col, c1: col })
  if (line?.t === "grand") apply(def.totalRow, { r0: row, r1: row, c0: L.c0, c1: L.c1 })
  return out
}
