import type { Border, CellStyle } from "../styles"
import { BUILT_IN_TABLE_STYLES } from "./built-in"
import { tableRows, type Table } from "./table"

/** One part of a table style (Excel's tableStyleElement): fill, bold, font colour and borders. */
export type TableStyleElement = {
  fill?: string
  bold?: true
  color?: string
  border?: Partial<Record<"left" | "top" | "bottom" | "right" | "vertical" | "horizontal", Border>>
  /** Rows (or columns) per stripe. */
  stripe?: number
}

export type TableStyleElementName =
  | "wholeTable"
  | "headerRow"
  | "totalRow"
  | "firstColumn"
  | "lastColumn"
  | "firstRowStripe"
  | "secondRowStripe"
  | "firstColumnStripe"
  | "secondColumnStripe"
  | "firstHeaderCell"
  | "lastHeaderCell"
  | "firstTotalCell"
  | "lastTotalCell"

export type TableStyleDef = Partial<Record<TableStyleElementName, TableStyleElement>>

/** The built-in style for a name ("TableStyleMedium2"), or null. */
export function tableStyleDef(name: string | null): TableStyleDef | null {
  if (!name) return null
  return BUILT_IN_TABLE_STYLES[name.replace(/^TableStyle/i, "")] ?? null
}

/** The built-in styles in Excel's gallery order, by group. */
export const TABLE_STYLE_GROUPS: { label: string; names: string[] }[] = [
  { label: "Light", names: Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`) },
  { label: "Medium", names: Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`) },
  { label: "Dark", names: Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`) },
]

/** What a table style puts on one cell: fill, bold, font colour and the four borders. */
export type TableCellLook = Pick<
  CellStyle,
  "fill" | "bold" | "fontColor" | "borderTop" | "borderBottom" | "borderLeft" | "borderRight"
>

type Box = { r0: number; r1: number; c0: number; c1: number }

/**
 * The look a table's style gives a cell of the table. Parts apply from the whole table up (Excel's order: whole
 * table, column stripes, row stripes, last column, first column, header row, total row, then the corner cells),
 * each over the ones before; a part's outer borders go on the edges of the area it covers, its inner ones between
 * its cells.
 */
export function tableCellLook(t: Table, row: number, col: number, def = tableStyleDef(t.style)): TableCellLook | null {
  if (!def) return null
  const r = t.range
  const parts = tableRows(t)
  const out: TableCellLook = {}
  const apply = (e: TableStyleElement | undefined, box: Box) => {
    if (!e) return
    if (e.fill) out.fill = e.fill
    if (e.bold) out.bold = true
    if (e.color) out.fontColor = e.color
    const b = e.border
    if (!b) return
    const side = (outer: Border | undefined, inner: Border | undefined, edge: boolean) => (edge ? outer : inner)
    const top = side(b.top, b.horizontal, row === box.r0)
    const bottom = side(b.bottom, b.horizontal, row === box.r1)
    const left = side(b.left, b.vertical, col === box.c0)
    const right = side(b.right, b.vertical, col === box.c1)
    if (top) out.borderTop = top
    if (bottom) out.borderBottom = bottom
    if (left) out.borderLeft = left
    if (right) out.borderRight = right
  }
  const [d0, d1] = parts.data
  const inData = row >= d0 && row <= d1
  apply(def.wholeTable, r)
  if (t.columnStripes && inData) {
    const first = def.firstColumnStripe
    const second = def.secondColumnStripe
    const a = first?.stripe ?? 1
    const b = second?.stripe ?? 1
    const at = (col - r.c0) % (a + b)
    const c0 = col - at
    if (at < a) apply(first, { r0: d0, r1: d1, c0, c1: c0 + a - 1 })
    else apply(second, { r0: d0, r1: d1, c0: c0 + a, c1: c0 + a + b - 1 })
  }
  if (t.rowStripes && inData) {
    const first = def.firstRowStripe
    const second = def.secondRowStripe
    const a = first?.stripe ?? 1
    const b = second?.stripe ?? 1
    const at = (row - d0) % (a + b)
    const r0 = row - at
    if (at < a) apply(first, { r0, r1: r0 + a - 1, c0: r.c0, c1: r.c1 })
    else apply(second, { r0: r0 + a, r1: r0 + a + b - 1, c0: r.c0, c1: r.c1 })
  }
  if (t.lastColumn && col === r.c1) apply(def.lastColumn, { r0: r.r0, r1: r.r1, c0: col, c1: col })
  if (t.firstColumn && col === r.c0) apply(def.firstColumn, { r0: r.r0, r1: r.r1, c0: col, c1: col })
  if (row === parts.header) {
    apply(def.headerRow, { r0: row, r1: row, c0: r.c0, c1: r.c1 })
    if (t.firstColumn && col === r.c0) apply(def.firstHeaderCell, { r0: row, r1: row, c0: col, c1: col })
    if (t.lastColumn && col === r.c1) apply(def.lastHeaderCell, { r0: row, r1: row, c0: col, c1: col })
  }
  if (row === parts.totals) {
    apply(def.totalRow, { r0: row, r1: row, c0: r.c0, c1: r.c1 })
    if (t.firstColumn && col === r.c0) apply(def.firstTotalCell, { r0: row, r1: row, c0: col, c1: col })
    if (t.lastColumn && col === r.c1) apply(def.lastTotalCell, { r0: row, r1: row, c0: col, c1: col })
  }
  return out
}

/** A cell's own format over its table's: what the cell sets itself wins, as in Excel. */
export function withTableLook(own: CellStyle, look: TableCellLook | null): CellStyle {
  if (!look) return own
  const out: CellStyle = { ...own }
  for (const k of Object.keys(look) as (keyof TableCellLook)[]) {
    if (out[k] === undefined) (out as Record<string, unknown>)[k] = look[k]
  }
  return out
}
