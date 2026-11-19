import type { CellValue } from "../values"
import { comparePivotValues, pivotFieldName, type PivotCell } from "./engine"
import {
  VALUES_FIELD,
  type PivotCache,
  type PivotField,
  type PivotSort,
  type PivotTable,
  type PivotValue,
} from "./model"

/**
 * The filter dropdowns of a PivotTable (the buttons on its Row Labels / Column Labels or field headers, and on its
 * report filters): which fields a dropdown covers, the items it lists, and the PivotTable with items ticked or
 * unticked, a field sorted, or a report filter's pick changed. Laying the cells out again is the caller's.
 */

/** What a filter button stands on: a header of the row or column area, or a report filter. */
export type PivotFilterTarget = {
  kind: "row" | "col" | "page"
  /** The fields the dropdown can work on (compact Row Labels covers every row field; "Select field" picks one). */
  fields: number[]
}

/** The filter button a cell of the PivotTable shows, if any. */
export function pivotFilterTarget(
  table: PivotTable,
  cache: PivotCache,
  cell: PivotCell | undefined,
): PivotFilterTarget | null {
  if (!cell) return null
  const real = (list: number[]) => list.filter((f) => f !== VALUES_FIELD)
  if (cell.role === "pageValue") {
    const page = table.pages[pageIndexOf(table, cell)]
    return page ? { kind: "page", fields: [page.field] } : null
  }
  if (cell.role === "rowHeader" || cell.role === "colHeader") {
    const row = cell.role === "rowHeader"
    const fields = real(row ? table.rows : table.cols)
    if (!fields.length) return null
    const header = String(cell.value ?? "")
    const caption = row ? (table.captions.rowHeader ?? "Row Labels") : (table.captions.colHeader ?? "Column Labels")
    // A field's own header (outline and tabular form) covers that field; the shared caption covers them all.
    const own = fields.filter((f) => header !== caption && pivotFieldName(table, cache, f) === header)
    return { kind: row ? "row" : "col", fields: own.length === 1 ? own : fields }
  }
  return null
}

/** The report filter a page value cell belongs to (they are laid out top down, in the order of `pages`). */
function pageIndexOf(table: PivotTable, cell: PivotCell): number {
  return cell.row - (table.at.row - table.pages.length - 1)
}

/** One item in a dropdown's list. */
export type PivotChoice = {
  /** The cache's shared item. */
  x: number
  label: string
  checked: boolean
}

export type PivotChoices = {
  field: number
  name: string
  /** What the sort commands say ("Sort A to Z" or "Smallest to Largest", "Oldest to Newest"). */
  kind: "text" | "number" | "date"
  sort: PivotSort
  items: PivotChoice[]
  /** Report filters: the one item picked, or null for (All) / several. */
  picked: number | null
  multipleItems: boolean
  /** Some item is unticked (or a report filter has a pick): the button shows a funnel. */
  filtered: boolean
}

/** The cache's shared items of a field (a column that keeps none: its distinct values, as the engine finds them). */
function sharedItems(cache: PivotCache, f: number): PivotValue[] {
  const kept = cache.fields[f]?.items
  if (kept) return kept
  const seen = new Set<string>()
  const out: PivotValue[] = []
  for (const rec of cache.records) {
    const v = rec[f] ?? null
    const k = v === null ? "" : typeof v === "object" ? `e${v.error}` : `${typeof v}:${String(v).toLowerCase()}`
    if (!seen.has(k)) {
      seen.add(k)
      out.push(v)
    }
  }
  return out
}

const textOf = (v: CellValue | null): string =>
  v === null || v === ""
    ? "(blank)"
    : typeof v === "object"
      ? v.error
      : typeof v === "boolean"
        ? v
          ? "TRUE"
          : "FALSE"
        : String(v)

/** Every item of a field once, in the field's order, with the cache's items it doesn't list yet after them. */
export function fullItems(field: PivotField, count: number): PivotField["items"] {
  const seen = new Set<number>()
  const items = field.items.filter((i) => i.x >= 0 && i.x < count && !seen.has(i.x) && seen.add(i.x))
  for (let x = 0; x < count; x++) if (!seen.has(x)) items.push({ x })
  return items
}

/**
 * The dropdown's list for a field: its items in the order the PivotTable shows them (sorted when the field is),
 * ticked unless hidden. `format` turns a value into the label shown (dates and numbers in the field's format).
 */
export function pivotChoices(
  table: PivotTable,
  cache: PivotCache,
  f: number,
  format: (v: CellValue, numFmt: string | undefined) => string = (v) => textOf(v),
): PivotChoices {
  const field = table.fields[f]
  const shared = sharedItems(cache, f)
  const items = fullItems(field, shared.length)
  const page = table.pages.find((p) => p.field === f)
  const picked = page && page.item !== null ? (items[page.item]?.x ?? null) : null
  const numFmt = field.numFmt ?? cache.fields[f]?.numFmt
  let list = items.map((item) => {
    const v = shared[item.x]
    const label = item.name ?? (v === null || v === "" ? "(blank)" : format(v, numFmt))
    return { x: item.x, label, checked: picked !== null ? item.x === picked : !item.hidden, value: v }
  })
  if (field.sort !== "manual") {
    const sign = field.sort === "ascending" ? 1 : -1
    list = [...list].sort((a, b) => {
      const blanks = Number(a.value === null || a.value === "") - Number(b.value === null || b.value === "")
      return blanks || sign * comparePivotValues(a.value, b.value)
    })
  }
  const present = shared.filter((v) => v !== null && v !== "")
  const kind = cache.fields[f]?.dates
    ? "date"
    : present.length && present.every((v) => typeof v === "number")
      ? "number"
      : "text"
  return {
    field: f,
    name: pivotFieldName(table, cache, f),
    kind,
    sort: field.sort,
    items: list.map(({ x, label, checked }) => ({ x, label, checked })),
    picked,
    multipleItems: field.multipleItems,
    filtered: picked !== null || items.some((i) => i.hidden),
  }
}

const withField = (table: PivotTable, f: number, field: PivotField): PivotTable => {
  const fields = [...table.fields]
  fields[f] = field
  return { ...table, fields }
}

/**
 * The table with only the items `shown` of field `f` ticked (null: all of them, Clear Filter). At least one item
 * stays: an empty set changes nothing. The same object when nothing changes.
 */
export function withShownItems(
  table: PivotTable,
  cache: PivotCache,
  f: number,
  shown: ReadonlySet<number> | null,
): PivotTable {
  const field = table.fields[f]
  if (!field || (shown && !shown.size)) return table
  const count = sharedItems(cache, f).length
  let changed = false
  const items = fullItems(field, count).map((item) => {
    const hide = shown !== null && !shown.has(item.x)
    if (Boolean(item.hidden) === hide) return item
    changed = true
    const next = { ...item }
    if (hide) next.hidden = true
    else delete next.hidden
    return next
  })
  // A report filter's single pick gives way to the ticked items.
  const pageAt = table.pages.findIndex((p) => p.field === f)
  let pages = table.pages
  if (pageAt >= 0 && table.pages[pageAt].item !== null) {
    pages = table.pages.map((p, i) => (i === pageAt ? { ...p, item: null } : p))
    changed = true
  }
  if (!changed) return table
  return { ...withField(table, f, { ...field, items }), pages }
}

/**
 * A report filter set to one item (`x`, a cache item) or to (All) (null), the way Excel's single-pick dropdown does:
 * every item ticked, the pick kept on the page field.
 */
export function withPagePick(table: PivotTable, cache: PivotCache, f: number, x: number | null): PivotTable {
  const pageAt = table.pages.findIndex((p) => p.field === f)
  const field = table.fields[f]
  if (pageAt < 0 || !field) return table
  const count = sharedItems(cache, f).length
  const items = fullItems(field, count).map((item) => {
    if (!item.hidden) return item
    const next = { ...item }
    delete next.hidden
    return next
  })
  const index = x === null ? null : items.findIndex((i) => i.x === x)
  if (index !== null && index < 0) return table
  const same =
    table.pages[pageAt].item === index &&
    !field.items.some((i) => i.hidden) &&
    field.items.length === items.length &&
    !field.multipleItems
  if (same) return table
  const pages = table.pages.map((p, i) => (i === pageAt ? { ...p, item: index } : p))
  return { ...withField(table, f, { ...field, items, multipleItems: false }), pages }
}

/** A report filter that lets several items be ticked ("Select Multiple Items"), with `shown` ticked. */
export function withPageItems(
  table: PivotTable,
  cache: PivotCache,
  f: number,
  shown: ReadonlySet<number> | null,
): PivotTable {
  const field = table.fields[f]
  if (!field || (shown && !shown.size)) return table
  const next = withShownItems(table, cache, f, shown)
  const nextField = next.fields[f]
  if (nextField.multipleItems) return next
  return withField(next, f, { ...nextField, multipleItems: true })
}

/**
 * The table with field `f` sorted (Sort A to Z / Z to A). Like Excel, the field's items are kept in the order shown,
 * so a report filter's pick and the rows written with the file point at the same items.
 */
export function withSort(table: PivotTable, cache: PivotCache, f: number, sort: PivotSort): PivotTable {
  const field = table.fields[f]
  if (!field || field.sort === sort) return table
  const shared = sharedItems(cache, f)
  let items = fullItems(field, shared.length)
  if (sort !== "manual") {
    const sign = sort === "ascending" ? 1 : -1
    items = [...items].sort((a, b) => {
      const va = a.name ?? shared[a.x]
      const vb = b.name ?? shared[b.x]
      const blanks = Number(va === null || va === "") - Number(vb === null || vb === "")
      return blanks || sign * comparePivotValues(va, vb)
    })
  }
  // A report filter's pick is an index into the items: follow the item it points at.
  const pageAt = table.pages.findIndex((p) => p.field === f)
  let pages = table.pages
  if (pageAt >= 0 && table.pages[pageAt].item !== null) {
    const before = fullItems(field, shared.length)[table.pages[pageAt].item!]
    const index = before ? items.findIndex((i) => i.x === before.x) : -1
    pages = table.pages.map((p, i) => (i === pageAt ? { ...p, item: index >= 0 ? index : null } : p))
  }
  return { ...withField(table, f, { ...field, sort, items }), pages }
}
