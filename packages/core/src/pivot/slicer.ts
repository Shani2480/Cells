// Slicers on PivotTables (PivotTable Analyze › Insert Slicer): one button per item of a cache field, filtering every
// PivotTable the slicer is connected to (Report Connections). Excel keeps the picks as hidden items of that field in
// each connected PivotTable, whether the field is on one of its axes or not (io's fixtures/excel/pivot-slicer.xlsx), so
// the PivotTables' own dropdowns and the slicer always agree. The buttons are listed as table slicers list them:
// sorted by value, "(blank)" last, items no record shows (given the PivotTable's other filters) after the rest.
import { displayText } from "../format"
import { compareSortText } from "../sort"
import type { SheetSlicer } from "../shapes"
import type { Sheet } from "../sheet"
import { slicerTarget, type SlicerItem } from "../slicers"
import { fullItems } from "./filter"
import { fieldShared, recordXs } from "./group"
import type { PivotCache, PivotTable, PivotValue } from "./model"

/** What a PivotTable slicer filters: a field of a PivotTable cache, in each of these PivotTables. */
export type PivotSlicerLink = {
  cacheId: number
  /** The cache field's name. */
  field: string
  /** The connected PivotTables, by sheet and name. */
  tables: { sheet: string; name: string }[]
}

export type PivotSlicerTarget = {
  cache: PivotCache
  field: number
  /** The connected PivotTables that still exist, with their sheets' indexes. */
  tables: { sheet: number; table: PivotTable }[]
}

/** A PivotTable slicer's cache, field and PivotTables (null: not a PivotTable slicer, or its field or tables gone). */
export function pivotSlicerTarget(
  sheets: readonly Sheet[],
  caches: readonly PivotCache[],
  slicer: Pick<SheetSlicer, "pivot">,
): PivotSlicerTarget | null {
  const link = slicer.pivot
  if (!link) return null
  const cache = caches.find((c) => c.id === link.cacheId)
  const field = cache ? cache.fields.findIndex((f) => f.name.toLowerCase() === link.field.toLowerCase()) : -1
  if (!cache || field < 0) return null
  const tables: PivotSlicerTarget["tables"] = []
  for (const t of link.tables) {
    const sheet = sheets.findIndex((s) => s.name.toLowerCase() === t.sheet.toLowerCase())
    const table = sheets[sheet]?.pivots.find((p) => p.name === t.name && p.cacheId === cache.id)
    if (table) tables.push({ sheet, table })
  }
  return tables.length ? { cache, field, tables } : null
}

const rank = (v: PivotValue) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3)
const isBlank = (v: PivotValue) => v === null || v === ""

/** The field's items in a slicer's order: by value (descending when the slicer says), "(blank)" last. Item indexes. */
export function pivotSlicerOrder(cache: PivotCache, f: number, descending = false): number[] {
  const values = fieldShared(cache, f)
  const xs = values.map((_, x) => x)
  const sorted = xs
    .filter((x) => !isBlank(values[x]))
    .sort((a, b) => {
      const va = values[a]
      const vb = values[b]
      const r = rank(va) - rank(vb)
      if (r) return r
      if (typeof va === "number" && typeof vb === "number") return va - vb
      return compareSortText(String(va).toLowerCase(), String(vb).toLowerCase())
    })
  if (descending) sorted.reverse()
  return [...sorted, ...xs.filter((x) => isBlank(values[x]))]
}

/** Which of field `f`'s items have records that pass the PivotTable's filters on its other fields. */
function itemsWithData(table: PivotTable, cache: PivotCache, f: number): Set<number> {
  const others: { xs: Int32Array; hidden: Set<number> }[] = []
  table.fields.forEach((field, g) => {
    if (g === f || !field || !cache.fields[g]) return
    const count = fieldShared(cache, g).length
    const items = fullItems(field, count)
    const page = table.pages.find((p) => p.field === g)
    const hidden = new Set<number>()
    if (page && page.item !== null) items.forEach((it, i) => i !== page.item && hidden.add(it.x))
    else for (const it of items) if (it.hidden) hidden.add(it.x)
    if (hidden.size) others.push({ xs: recordXs(cache, g), hidden })
  })
  const own = recordXs(cache, f)
  const out = new Set<number>()
  next: for (let r = 0; r < cache.records.length; r++) {
    for (const o of others) if (o.hidden.has(o.xs[r])) continue next
    if (own[r] >= 0) out.add(own[r])
  }
  return out
}

/** A PivotTable slicer's buttons: keys are `x:<item index>`; picked unless hidden in the first connected PivotTable. */
export function pivotSlicerItems(target: PivotSlicerTarget, slicer: SheetSlicer): SlicerItem[] {
  const { cache, field: f } = target
  const first = target.tables[0].table
  const values = fieldShared(cache, f)
  const hidden = new Set(first.fields[f]?.items.filter((i) => i.hidden).map((i) => i.x) ?? [])
  const page = first.pages.find((p) => p.field === f)
  const pick = page && page.item !== null ? first.fields[f]?.items[page.item]?.x : undefined
  const data = slicer.crossFilter === "none" ? null : itemsWithData(first, cache, f)
  const fmt = first.fields[f]?.numFmt ?? cache.fields[f].numFmt ?? "General"
  let items = pivotSlicerOrder(cache, f, slicer.sort === "descending").map((x): SlicerItem => {
    const v = values[x]
    return {
      key: `x:${x}`,
      text: isBlank(v) ? "(blank)" : typeof v === "number" ? displayText(v, fmt) : displayText(v, "General"),
      selected: pick !== undefined ? x === pick : !hidden.has(x),
      data: data ? data.has(x) : true,
    }
  })
  if (slicer.hideNoData) items = items.filter((e) => e.data)
  else if (slicer.crossFilter === "showItemsWithDataAtTop")
    items = [...items.filter((e) => e.data), ...items.filter((e) => !e.data)]
  return items
}

/** The item indexes a set of picked button keys stands for (null: every button picked, no filter). */
export function pivotSlicerShown(items: readonly SlicerItem[], picked: ReadonlySet<string>): Set<number> | null {
  if (items.every((i) => picked.has(i.key))) return null
  return new Set([...picked].filter((k) => k.startsWith("x:")).map((k) => Number(k.slice(2))))
}

/** Is the slicer filtering (an item hidden in its first PivotTable)? */
export function pivotSlicerFiltered(target: PivotSlicerTarget): boolean {
  const first = target.tables[0].table
  const page = first.pages.find((p) => p.field === target.field)
  return Boolean(first.fields[target.field]?.items.some((i) => i.hidden) || (page && page.item !== null))
}

/** A sheet or PivotTable renamed: PivotTable slicers follow (the ops of the same undo step give the new objects). */
export function pivotSlicerRenamed(
  link: PivotSlicerLink,
  from: { sheet: string; name?: string },
  to: { sheet: string; name?: string },
): PivotSlicerLink | null {
  let changed = false
  const tables = link.tables.map((t) => {
    if (t.sheet.toLowerCase() !== from.sheet.toLowerCase()) return t
    if (from.name !== undefined && t.name !== from.name) return t
    changed = true
    return { sheet: to.sheet, name: from.name !== undefined ? (to.name ?? t.name) : t.name }
  })
  return changed ? { ...link, tables } : null
}

/** Does a slicer still have something to filter (its table column, or a PivotTable)? */
export function slicerConnected(
  sheets: readonly Sheet[],
  caches: readonly PivotCache[],
  slicer: Pick<SheetSlicer, "table" | "column" | "pivot">,
): boolean {
  return slicer.pivot ? pivotSlicerTarget(sheets, caches, slicer) !== null : slicerTarget(sheets, slicer) !== null
}
