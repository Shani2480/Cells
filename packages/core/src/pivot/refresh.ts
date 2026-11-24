import { fieldShared, inRecords, refreshedDiscrete } from "./group"
import type { Range } from "../address"
import { chartRefRange } from "../chart"
import { quoteSheet } from "../formula/print"
import type { SheetList } from "../names"
import { isDateFormat } from "../numfmt"
import type { CellWrite, Op } from "../ops"
import type { Sheet } from "../sheet"
import type { StyleTable } from "../styles"
import { findTable } from "../table/table"
import { pivotLayout, type PivotLayout } from "./engine"
import {
  newPivotField,
  VALUES_FIELD,
  type PivotCache,
  type PivotCacheField,
  type PivotTable,
  type PivotValue,
} from "./model"
import { copyPivotTable, pivotArea } from "./sheet"

/**
 * Refresh (PivotTable Analyze › Refresh, Data › Refresh All): the cache reads its source again, each PivotTable on
 * it keeps its settings for the fields and items that are still there, and its cells are laid out anew.
 */

export const PIVOT_SOURCE_MESSAGE = "Reference isn't valid."
export const PIVOT_FIELD_NAME_MESSAGE =
  "The PivotTable field name is not valid. To create a PivotTable report, you must use data that is organized as a " +
  "list with labeled columns. If you are changing the name of a PivotTable field, you must type a new name for the field."

/** The cells a cache reads: its range, or the table or defined name it names (headers in the first row). */
export function pivotSourceRange(
  sheets: readonly Sheet[],
  source: PivotCache["source"],
): { sheet: Sheet; range: Range } | null {
  if (source.kind === "range") return chartRefRange(sheets, `${quoteSheet(source.sheet)}!${source.ref}`)
  const table = findTable(sheets, source.name)
  if (table) return { sheet: sheets[table.sheet], range: { ...table.table.range } }
  const def = (sheets as SheetList).names?.get(source.name)
  return def ? chartRefRange(sheets, def.ref) : null
}

/** The source's columns and rows, as a cache keeps them; a message when it can't be read. */
export function readPivotSource(
  sheets: readonly Sheet[],
  styles: StyleTable,
  source: PivotCache["source"],
): { fields: Omit<PivotCacheField, "items">[]; records: PivotValue[][] } | string {
  const at = pivotSourceRange(sheets, source)
  if (!at) return PIVOT_SOURCE_MESSAGE
  const { sheet } = at
  // A whole-column reference stops at the data in use.
  const used = sheet.cells.usedBounds()
  const range = { ...at.range }
  if (used) range.r1 = Math.min(range.r1, used.r1)
  if (range.r1 < range.r0) return PIVOT_SOURCE_MESSAGE
  const fields: Omit<PivotCacheField, "items">[] = []
  for (let c = range.c0; c <= range.c1; c++) {
    const v = sheet.cells.get(range.r0, c)?.value
    const name = v === undefined || v === null || v === "" ? "" : typeof v === "object" ? v.error : String(v)
    if (!name) return PIVOT_FIELD_NAME_MESSAGE
    fields.push({ name })
  }
  const records: PivotValue[][] = []
  const dates = fields.map(() => ({ dates: 0, numbers: 0, fmt: undefined as string | undefined }))
  for (let r = range.r0 + 1; r <= range.r1; r++) {
    const rec: PivotValue[] = []
    for (let c = range.c0; c <= range.c1; c++) {
      const v = sheet.cells.get(r, c)?.value
      rec.push(v === undefined || v === "" ? null : v)
      if (typeof v === "number") {
        const d = dates[c - range.c0]
        const fmt = styles.get(sheet.styleAt(r, c)).numberFormat
        d.numbers++
        if (fmt && isDateFormat(fmt)) d.dates++
        d.fmt ??= fmt && fmt !== "General" ? fmt : undefined
      }
    }
    records.push(rec)
  }
  fields.forEach((f, i) => {
    const d = dates[i]
    if (d.fmt) f.numFmt = d.fmt
    if (d.numbers && d.dates === d.numbers) f.dates = true
  })
  return { fields, records }
}

const keyOf = (v: PivotValue): string =>
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
 * The cache after reading its source again: a field keeps its shared items that are still there, in their order,
 * then the new ones as first met (so the PivotTables' items keep their places).
 */
export function refreshedCache(old: PivotCache, data: Exclude<ReturnType<typeof readPivotSource>, string>): PivotCache {
  const fields = data.fields.map((f, i): PivotCacheField => {
    const before = old.fields.find((o) => o.name.toLowerCase() === f.name.toLowerCase())
    // A column without shared items keeps none (the engine finds its values when it needs them).
    if (!before?.items) return { ...f, items: null }
    // Calculated items stay, after the others.
    const b = old.fields.indexOf(before)
    const calc = (old.calcItems ?? []).filter((ci) => ci.field === b).map((ci) => before.items![ci.x])
    const calcKeys = new Set(calc.map(keyOf))
    const present = new Map<string, PivotValue>()
    for (const r of data.records) {
      const k = keyOf(r[i] ?? null)
      if (!present.has(k) && !calcKeys.has(k)) present.set(k, r[i] ?? null)
    }
    const items: PivotValue[] = []
    const seen = new Set<string>()
    for (const v of before.items) {
      const k = keyOf(v)
      if (present.has(k) && !seen.has(k)) {
        seen.add(k)
        items.push(present.get(k)!)
      }
    }
    for (const [k, v] of present) if (!seen.has(k)) items.push(v)
    return { ...f, items: [...items, ...calc] }
  })
  // The calculated items follow their fields and shared items.
  const calcItems = (old.calcItems ?? []).flatMap((ci) => {
    const name = old.fields[ci.field]?.name.toLowerCase()
    const field = fields.findIndex((f) => f.name.toLowerCase() === name)
    const x =
      field >= 0
        ? (fields[field].items?.findIndex((v) => keyOf(v) === keyOf(old.fields[ci.field].items![ci.x])) ?? -1)
        : -1
    return x >= 0 ? [{ ...ci, field, x }] : []
  })
  const indexOf = (name: string) => fields.findIndex((f) => f.name.toLowerCase() === name.toLowerCase())
  // Number ranges stay on their column, with the same start, end and labels (new values outside go to < and >).
  old.fields.forEach((f, i) => {
    if (f.group?.kind !== "range" || f.group.base !== i) return
    const j = indexOf(f.name)
    if (j >= 0) fields[j] = { ...fields[j], group: { ...f.group, base: j } }
  })
  // Calculated fields and group fields stay, after the source's columns, in their order (one named like a new column
  // gives way to it; a group field whose base is gone goes too).
  const taken = new Set(fields.map((f) => f.name.toLowerCase()))
  old.fields.forEach((f, i) => {
    if (inRecords(f, i) || taken.has(f.name.toLowerCase())) return
    if (!f.group) {
      fields.push({ ...f })
      return
    }
    const base = indexOf(old.fields[f.group.base]?.name ?? "")
    if (base < 0) return
    if (f.group.kind === "range") fields.push({ ...f, group: { ...f.group, base } })
    else {
      const sofar: PivotCache = { ...old, fields: [...fields], records: data.records }
      const oldBase = fieldShared(old, f.group.base)
      fields.push({ ...f, group: refreshedDiscrete(f.group, oldBase, fieldShared(sofar, base), base) })
    }
  })
  // A discrete group points at its base's shared items: keep them.
  fields.forEach((f) => {
    if (f.group?.kind !== "discrete") return
    const base = f.group.base
    if (!fields[base].items && !fields[base].group && inRecords(fields[base], base))
      fields[base] = { ...fields[base], items: fieldShared({ ...old, fields, records: data.records }, base) }
  })
  const next: PivotCache = { ...old, fields, records: data.records }
  if (calcItems.length) next.calcItems = calcItems
  else delete next.calcItems
  return next
}

/** A field's item values (its group items, the cache's shared items, or its column's distinct values). */
const itemValues = fieldShared

/** The PivotTable on the refreshed cache: fields matched by name, items by value, the rest as Excel makes them. */
export function refreshedTable(table: PivotTable, old: PivotCache, next: PivotCache): PivotTable {
  const t = copyPivotTable(table)
  const fieldMap = new Map<number, number>()
  old.fields.forEach((f, i) => {
    const j = next.fields.findIndex((n) => n.name.toLowerCase() === f.name.toLowerCase())
    if (j >= 0) fieldMap.set(i, j)
  })
  const pagePos = new Map<number, number | null>()
  t.fields = next.fields.map((_f, j) => {
    const i = [...fieldMap].find(([, to]) => to === j)?.[0]
    const before = i !== undefined ? table.fields[i] : undefined
    const nextValues = itemValues(next, j)
    if (!before) {
      const field = newPivotField(next.fields[j].items || next.fields[j].group ? nextValues.length : 0, t)
      // A calculated field is only ever a value: Excel gives it no subtotals.
      if (next.fields[j].formula !== undefined) field.subtotals = []
      return field
    }
    const oldValues = itemValues(old, i!)
    const index = new Map<string, number>()
    nextValues.forEach((v, x) => index.set(keyOf(v), x))
    const items: PivotTable["fields"][number]["items"] = []
    const posMap = new Map<number, number>()
    before.items.forEach((it, pos) => {
      const x = index.get(keyOf(oldValues[it.x] ?? null))
      if (x === undefined) return
      posMap.set(pos, items.length)
      items.push({ ...it, x })
    })
    // New items go after the listed ones (shared items with no data stay out of a field that lists none).
    if (next.fields[j].items && before.items.length) {
      const listed = new Set(items.map((it) => it.x))
      nextValues.forEach((_v, x) => listed.has(x) || items.push({ x }))
    }
    const page = table.pages.find((p) => p.field === i)
    if (page) pagePos.set(j, page.item === null ? null : (posMap.get(page.item) ?? null))
    return { ...before, items, subtotals: [...before.subtotals] }
  })
  const map = (f: number) => (f === VALUES_FIELD ? f : (fieldMap.get(f) ?? -1))
  t.values = table.values
    .filter((v) => fieldMap.has(v.field))
    .map((v) => {
      const moved = { ...v, field: fieldMap.get(v.field)! }
      // Show Values As' base field follows its column; gone, the value shows plainly again.
      if (v.baseField !== undefined) {
        const base = fieldMap.get(v.baseField)
        if (base !== undefined) {
          moved.baseField = base
          if (typeof v.baseItem === "number") {
            // The base item by its value (-1, shown as #N/A, when it is gone).
            const value = itemValues(old, v.baseField)[v.baseItem]
            const k = keyOf(value ?? null)
            moved.baseItem = itemValues(next, base).findIndex((n) => keyOf(n) === k)
          }
        } else {
          delete moved.showAs
          delete moved.baseField
          delete moved.baseItem
        }
      }
      return moved
    })
  t.rows = table.rows.map(map).filter((f) => f !== -1)
  t.cols = table.cols.map(map).filter((f) => f !== -1)
  if (t.values.length < 2) {
    t.rows = t.rows.filter((f) => f !== VALUES_FIELD)
    t.cols = t.cols.filter((f) => f !== VALUES_FIELD)
  }
  t.pages = table.pages
    .filter((p) => fieldMap.has(p.field))
    .map((p) => ({ field: fieldMap.get(p.field)!, item: pagePos.get(fieldMap.get(p.field)!) ?? null }))
  t.fields.forEach((f, j) => {
    f.dataField = t.values.some((v) => v.field === j)
    if (f.axis === "row" && !t.rows.includes(j)) f.axis = null
    if (f.axis === "col" && !t.cols.includes(j)) f.axis = null
    if (f.axis === "page" && !t.pages.some((p) => p.field === j)) f.axis = null
  })
  return t
}

export type PivotCellsPlan = {
  /** setCells for the old cells that go and the new ones (the caller puts `table` in the sheet's list). */
  ops: Op[]
  /** The PivotTable with its new location. */
  table: PivotTable
  layout: PivotLayout
  /** Cells outside the PivotTable's old place that hold data and would be written over. */
  overwrites: Range | null
}

/**
 * The cells a PivotTable shows, written into its sheet: its old cells cleared (formats kept), the new layout's
 * values written with the value fields' number formats.
 */
export function pivotCellsPlan(
  sheets: readonly Sheet[],
  styles: StyleTable,
  sheetIndex: number,
  before: PivotTable | null,
  table: PivotTable,
  cache: PivotCache,
): PivotCellsPlan {
  const sheet = sheets[sheetIndex]
  const layout = pivotLayout(table, cache)
  const t = copyPivotTable(table)
  t.location = { ...layout.range }
  t.at = { row: layout.range.r0, col: layout.range.c0 }
  const writes = new Map<string, CellWrite>()
  /**
   * A cell's style with the indent the layout gives it (Excel saves compact labels' indent in the cells' formats, so
   * a label moving to another line takes its indent along, and what's left behind loses it).
   */
  const indented = (style: number, indent: number | undefined) => {
    const now = styles.get(style)
    if ((now.indent ?? 0) === (indent ?? 0)) return style
    const next = { ...now }
    if (indent) next.indent = indent
    else delete next.indent
    return styles.intern(next)
  }
  const keep = (row: number, col: number) => {
    const style = indented(sheet.cells.style(row, col), undefined)
    return style ? { style } : undefined
  }
  if (before) {
    const a = pivotArea(before)
    for (let r = a.r0; r <= a.r1; r++)
      for (let c = a.c0; c <= a.c1; c++)
        if (sheet.cells.get(r, c)?.value !== undefined) writes.set(`${r},${c}`, { row: r, col: c, data: keep(r, c) })
  }
  const old = before ? pivotArea(before) : null
  const area = pivotArea(t)
  let overwrites: Range | null = null
  for (let r = area.r0; r <= area.r1; r++)
    for (let c = area.c0; c <= area.c1; c++) {
      const inside = old && r >= old.r0 && r <= old.r1 && c >= old.c0 && c <= old.c1
      if (!inside && sheet.cells.get(r, c)?.value !== undefined) {
        overwrites = overwrites
          ? {
              r0: Math.min(overwrites.r0, r),
              c0: Math.min(overwrites.c0, c),
              r1: Math.max(overwrites.r1, r),
              c1: Math.max(overwrites.c1, c),
            }
          : { r0: r, c0: c, r1: r, c1: c }
      }
    }
  for (const cell of layout.cells) {
    if (cell.value === undefined) continue
    let style = indented(sheet.cells.style(cell.row, cell.col), cell.indent)
    if (cell.numFmt) style = styles.intern({ ...styles.get(style), numberFormat: cell.numFmt })
    writes.set(`${cell.row},${cell.col}`, {
      row: cell.row,
      col: cell.col,
      data: style ? { value: cell.value, style } : { value: cell.value },
    })
  }
  const ops: Op[] = [{ kind: "setCells", sheet: sheetIndex, cells: [...writes.values()] }]
  return { ops, table: t, layout, overwrites }
}

/** Excel's question before a PivotTable grows over cells with data. */
export const pivotOverwriteMessage = (book: string, sheet: string) =>
  `There's already data in [${book}]${sheet}. Do you want to replace it?`

/** A new cache from its source's data: shared items for columns with text (number and date columns keep none). */
export function pivotCacheFrom(
  id: number,
  source: PivotCache["source"],
  data: Exclude<ReturnType<typeof readPivotSource>, string>,
): PivotCache {
  const fields = data.fields.map((f, i): PivotCacheField => {
    const numeric = data.records.every((r) => r[i] === null || typeof r[i] === "number")
    if (numeric) return { ...f, items: null }
    const items: PivotValue[] = []
    const seen = new Set<string>()
    for (const r of data.records) {
      const k = keyOf(r[i] ?? null)
      if (!seen.has(k)) {
        seen.add(k)
        items.push(r[i] ?? null)
      }
    }
    return { ...f, items }
  })
  return { id, source, fields, records: data.records, refreshOnLoad: false }
}
