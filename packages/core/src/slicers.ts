import type { Range } from "./address"
import type { CellLook } from "./conditional"
import { filterColumnCells, filterVisibility, type AutoFilter, type ColumnFilter, type DateGroup } from "./filter"
import { serialToDate } from "./formula/dates"
import { isDateFormat } from "./numfmt"
import type { Op } from "./ops"
import { isChart, type SheetSlicer } from "./shapes"
import type { Sheet } from "./sheet"
import { compareSortText } from "./sort"
import type { StyleTable } from "./styles"
import type { Table } from "./table/table"
import { applyTint, type Theme } from "./theme"

/**
 * Slicers for tables (Insert › Slicer): one button per value of a table column. Which buttons are picked is the
 * table's own filter on that column, as Excel keeps it, so the Filter Button's dropdown and the slicer always agree.
 */

/** Excel's slicer styles, in the gallery's order (Light, Other, Dark). */
export const SLICER_STYLES = [
  "SlicerStyleLight1",
  "SlicerStyleLight2",
  "SlicerStyleLight3",
  "SlicerStyleLight4",
  "SlicerStyleLight5",
  "SlicerStyleLight6",
  "SlicerStyleOther1",
  "SlicerStyleOther2",
  "SlicerStyleDark1",
  "SlicerStyleDark2",
  "SlicerStyleDark3",
  "SlicerStyleDark4",
  "SlicerStyleDark5",
  "SlicerStyleDark6",
] as const

export const DEFAULT_SLICER_STYLE = "SlicerStyleLight1"

/** A new slicer's size and button height, as Excel makes it (2 × 2.53 in; 0.27 in buttons). */
export const SLICER_SIZE = { w: 192, h: 243 }
export const SLICER_BUTTON_HEIGHT = 27

/** One button: the text it shows, whether it's picked, and whether the table's other filters leave rows with it. */
export type SlicerItem = { key: string; text: string; selected: boolean; data: boolean }

/** The table a slicer filters, its sheet and the column's index in the sheet; null when either is gone. */
export function slicerTarget(
  sheets: readonly Sheet[],
  slicer: Pick<SheetSlicer, "table" | "column">,
): { sheet: number; table: Table; col: number } | null {
  for (let i = 0; i < sheets.length; i++)
    for (const table of sheets[i].tables) {
      if (table.id !== slicer.table) continue
      const at = table.columns.findIndex((c) => c.name.toLowerCase() === slicer.column.toLowerCase())
      return at < 0 ? null : { sheet: i, table, col: table.range.c0 + at }
    }
  return null
}

/** A table's columns renamed (its header typed over): its slicers follow, in the same undo step. */
export function slicerColumnOps(
  sheets: readonly Sheet[],
  tableId: number,
  renames: readonly { from: string; to: string }[],
): Op[] {
  const ops: Op[] = []
  sheets.forEach((sheet, i) => {
    let changed = false
    const objects = sheet.objects.map((o) => {
      if (isChart(o) || o.kind !== "slicer" || o.table !== tableId) return o
      const r = renames.find((x) => x.from.toLowerCase() === o.column.toLowerCase())
      if (!r) return o
      changed = true
      return { ...o, column: r.to }
    })
    if (changed) ops.push({ kind: "objects", sheet: i, objects })
  })
  return ops
}

/** The range a table's filter covers: the header and the data rows (not the total row). */
export function tableFilterRange(table: Table): Range {
  return { ...table.range, r1: table.range.r1 - (table.totals ? 1 : 0) }
}

const norm = (text: string) => text.trim().toLowerCase()
const rank = (v: unknown) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3)

/**
 * A slicer's buttons: each value of the column once (dates by day, "(blank)" for empty cells, last), sorted as the
 * slicer says, with the items the other filters leave no rows for after the rest (or in place, or not told apart).
 */
export function slicerItems(
  sheets: readonly Sheet[],
  styles: StyleTable,
  slicer: SheetSlicer,
  today: number,
  look?: CellLook,
): SlicerItem[] {
  const target = slicerTarget(sheets, slicer)
  if (!target) return []
  const sheet = sheets[target.sheet]
  const range = target.table.autoFilter?.range ?? tableFilterRange(target.table)
  const filter: AutoFilter = target.table.autoFilter ?? { range, columns: {} }
  const offset = target.col - range.c0
  const own = filter.columns[offset]
  const others: AutoFilter = { range, columns: { ...filter.columns } }
  delete others.columns[offset]
  const left = filterVisibility(sheet, styles, others, today, look)
  const passing = own ? filterVisibility(sheet, styles, { range, columns: { [offset]: own } }, today, look) : null
  type Entry = SlicerItem & { value: unknown; order: number }
  const entries = new Map<string, Entry>()
  filterColumnCells(sheet, styles, range, target.col).forEach((cell, i) => {
    const value = cell.data?.value
    const blank = value === undefined || value === null || value === ""
    const date = typeof value === "number" && cell.code && isDateFormat(cell.code)
    const key = blank ? "" : date ? `d:${Math.floor(value)}` : norm(cell.text)
    let e = entries.get(key)
    if (!e) {
      e = { key, text: blank ? "(blank)" : cell.text, selected: false, data: false, value, order: entries.size }
      entries.set(key, e)
    }
    e.selected ||= passing ? passing[i] : true
    e.data ||= left[i]
  })
  const list = [...entries.values()]
  const blank = list.filter((e) => e.key === "")
  const sorted = list
    .filter((e) => e.key !== "")
    .sort((a, b) => {
      const r = rank(a.value) - rank(b.value)
      if (r) return r
      if (typeof a.value === "number" && typeof b.value === "number") return a.value - b.value
      return compareSortText(norm(a.text), norm(b.text))
    })
  if (slicer.sort === "descending") sorted.reverse()
  let items = [...sorted, ...blank]
  if (slicer.crossFilter === "none") items = items.map((e) => ({ ...e, data: true }))
  else if (slicer.hideNoData) items = items.filter((e) => e.data)
  else if (slicer.crossFilter === "showItemsWithDataAtTop")
    items = [...items.filter((e) => e.data), ...items.filter((e) => !e.data)]
  return items.map(({ key, text, selected, data }) => ({ key, text, selected, data }))
}

/** The column filter that shows the picked buttons' rows (null when every button is picked: no filter). */
export function slicerFilter(items: readonly SlicerItem[], picked: ReadonlySet<string>): ColumnFilter | null {
  if (items.every((i) => picked.has(i.key))) return null
  const values: string[] = []
  const dates: DateGroup[] = []
  let blanks = false
  for (const item of items) {
    if (!picked.has(item.key)) continue
    if (item.key === "") blanks = true
    else if (item.key.startsWith("d:")) {
      const d = serialToDate(Number(item.key.slice(2)))
      dates.push({ year: d.year, month: d.month, day: d.day })
    } else values.push(item.text)
  }
  return { kind: "values", values, ...(blanks ? { blanks: true } : {}), ...(dates.length ? { dates } : {}) }
}

// ── Styles ─────────────────────────────────────────────────────────────────────────────────────────────────────

/** How one kind of button looks. */
export type SlicerButtonLook = { fill: string; border: string; text: string }

/** A style's colours (as "#RRGGBB" under a theme), sampled from Excel 365. */
export type SlicerLook = {
  frame: string
  background: string
  heading: string
  /** The line under the header (null: none). */
  divider: string | null
  selected: SlicerButtonLook
  unselected: SlicerButtonLook
  selectedNoData: SlicerButtonLook
  unselectedNoData: SlicerButtonLook
}

const ACCENT_SLOTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const

/** A style's colours under `theme` (an unknown style looks like Light 1). */
export function slicerLook(style: string, theme: Theme): SlicerLook {
  const m = /^SlicerStyle(Light|Dark|Other)(\d)$/.exec(style)
  const family = m?.[1] ?? "Light"
  const n = Math.min(6, Math.max(1, Number(m?.[2] ?? 1)))
  const accent = (theme.colors[ACCENT_SLOTS[n - 1]] ?? "156082").toUpperCase()
  const hex = (tint: number) => "#" + (tint ? applyTint(accent, tint) : accent)
  const gray = "#7F7F7F"
  if (family === "Other" && n === 1)
    return {
      frame: "#7F7F7F",
      background: "#FFFFFF",
      heading: "#000000",
      divider: "#999999",
      selected: { fill: "#BEBEBE", border: "#999999", text: "#000000" },
      unselected: { fill: "#FFFFFF", border: "#CCCCCC", text: "#000000" },
      selectedNoData: { fill: "#D8D8D8", border: "#CCCCCC", text: "#000000" },
      unselectedNoData: { fill: "#FFFFFF", border: "#DBDBDB", text: "#808080" },
    }
  if (family === "Other")
    return {
      frame: "#4F81BD",
      background: "#FFFFFF",
      heading: "#000000",
      divider: null,
      selected: { fill: "#B6CCEA", border: "#999999", text: "#000000" },
      unselected: { fill: "#F0F1F4", border: "#CCCCCC", text: "#000000" },
      selectedNoData: { fill: "#DDE9F8", border: "#CCCCCC", text: "#828282" },
      unselectedNoData: { fill: "#F8F9FA", border: "#E0E0E0", text: "#828282" },
    }
  if (family === "Dark")
    return {
      frame: hex(0),
      background: "#FFFFFF",
      heading: "#000000",
      divider: null,
      selected: { fill: hex(0), border: hex(0), text: "#FFFFFF" },
      unselected: { fill: "#C0C0C0", border: "#C0C0C0", text: "#000000" },
      selectedNoData: { fill: hex(0.6), border: hex(0.6), text: hex(-0.25) },
      unselectedNoData: { fill: "#DEDEDE", border: "#DBDBDB", text: "#959595" },
    }
  return {
    frame: hex(0),
    background: "#FFFFFF",
    heading: "#000000",
    divider: hex(0),
    selected: { fill: hex(0.6), border: "#999999", text: "#000000" },
    unselected: { fill: "#FFFFFF", border: "#CCCCCC", text: "#000000" },
    selectedNoData: { fill: hex(0.8), border: "#CCCCCC", text: gray },
    unselectedNoData: { fill: "#FFFFFF", border: "#DBDBDB", text: gray },
  }
}

// ── Layout ─────────────────────────────────────────────────────────────────────────────────────────────────────

export type Box = { x: number; y: number; w: number; h: number }

/** Where things sit in a slicer of `w` × `h` (px at 100%), as Excel lays it out. */
export type SlicerLayout = {
  header: Box | null
  /** The Multi-Select and Clear Filter buttons in the header. */
  multi: Box | null
  clear: Box | null
  divider: number | null
  /** Each shown item's button, from `first` on (the rest are scrolled out of view). */
  buttons: Box[]
  first: number
  /** More buttons than fit: Excel shows a scroll bar. */
  scroll: { total: number; shown: number } | null
}

const PAD = 8
const HEADER = 29
const GAP = 3

export function slicerLayout(
  slicer: Pick<SheetSlicer, "header" | "columns" | "buttonHeight" | "startItem">,
  w: number,
  h: number,
  count: number,
): SlicerLayout {
  const header = slicer.header ? { x: 1, y: 1, w: w - 2, h: HEADER - 1 } : null
  const top = slicer.header ? HEADER + 5 : PAD + 1
  const left = PAD + 1
  const columns = Math.max(1, slicer.columns)
  const bh = Math.max(4, slicer.buttonHeight)
  const rows = Math.ceil(count / columns)
  const fit = Math.max(1, Math.floor((h - top - PAD + GAP) / (bh + GAP)))
  const scroll = rows > fit ? { total: rows, shown: fit } : null
  const inner = w - left - PAD - 1 - (scroll ? 12 : 0)
  const bw = Math.max(4, (inner - GAP * (columns - 1)) / columns)
  const firstRow = Math.max(0, Math.min(rows - fit, Math.floor((slicer.startItem ?? 0) / columns)))
  const buttons: Box[] = []
  for (let i = firstRow * columns; i < count && Math.floor(i / columns) < firstRow + fit; i++) {
    const r = Math.floor(i / columns) - firstRow
    const c = i % columns
    buttons.push({ x: left + c * (bw + GAP), y: top + r * (bh + GAP), w: bw, h: bh })
  }
  return {
    header,
    multi: header ? { x: w - 60, y: 6, w: 18, h: 18 } : null,
    clear: header ? { x: w - 32, y: 6, w: 20, h: 18 } : null,
    divider: slicer.header ? HEADER : null,
    buttons,
    first: firstRow * columns,
    scroll,
  }
}
