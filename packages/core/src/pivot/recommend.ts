// Recommended PivotTables (Insert › Recommended PivotTables, Quick Analysis' Tables tab): a short list of
// PivotTables worth making from a range, named as Excel names its suggestions ("Sum of Amount by Region"). Excel's
// own picks come from a model of its own; these follow its kinds of suggestion: one category field on the rows with
// a number summed, dates grouped by months (and years when the data spans several), two categories crossed, and a
// count when the data has no numbers. Categories with few items come first; a column that is all different values
// (an ID, a name per row) is never a category, nor is a column of numbers.
import { fieldShared, groupKind, derivedDateGroups, withRangeGroup, inRecords } from "./group"
import { newPivotTable, withFieldIn } from "./make"
import { withFieldFilter, withSort, withValueSort } from "./filter"
import { withAllCleared, withValueSettings } from "./settings"
import { serialToDate } from "../formula/dates"
import type { CellAddress } from "../address"
import type { PivotCache, PivotFn, PivotTable } from "./model"

export type PivotRecipe = {
  /** As Excel titles it: "Sum of Amount by Region", "Count of Product by Region and Channel". */
  title: string
  rows: number[]
  cols: number[]
  /** The fields summed (counted when they hold text). */
  values: number[]
  /** How each of `values` is summarized, when not the usual Sum (numbers) or Count (text). */
  fns?: (PivotFn | undefined)[]
  /** Report filters (all items shown). */
  pages?: number[]
  /** A date field on the rows or columns, grouped by these parts (largest last, as Grouping lists them). */
  dates?: { field: number; parts: DatePart[] }
  /** One row or column field sorted: by its labels, or by its totals of value `by` (an index into `values`). */
  sort?: { field: number; by: "label" | number; order: "ascending" | "descending" }
  /** Only the `count` items of a row or column field with the biggest (smallest: `bottom`) totals of value `measure`. */
  top?: { field: number; count: number; bottom?: boolean; measure: number }
}

/** The parts a date field can be grouped by in a recipe. */
export type DatePart = "days" | "months" | "quarters" | "years"

const MAX_ITEMS = 40

/** Up to `limit` PivotTables worth making from the cache's data, best first. */
export function recommendedPivots(cache: PivotCache, limit = 10): PivotRecipe[] {
  const count = cache.records.length
  if (!count) return []
  const plain = cache.fields.map((field, f) => ({ field, f })).filter(({ field, f }) => inRecords(field, f))
  const kind = (f: number) => {
    const values = fieldShared(cache, f).filter((v) => v !== null && v !== "")
    if (!values.length) return "empty"
    if (cache.fields[f].dates || groupKind(cache, f) === "dates") return "date"
    if (values.every((v) => typeof v === "number")) return "number"
    return "text"
  }
  const kinds = new Map(plain.map(({ f }) => [f, kind(f)]))
  const items = (f: number) => fieldShared(cache, f).length
  // Categories: text with some repeats, fewest items first; dates when they span more than a month.
  const categories = plain
    .filter(({ f }) => kinds.get(f) === "text" && items(f) >= 2 && items(f) <= MAX_ITEMS && items(f) < count)
    .map(({ f }) => f)
    .sort((a, b) => items(a) - items(b))
  const numbers = plain.filter(({ f }) => kinds.get(f) === "number").map(({ f }) => f)
  const dates = plain.filter(({ f }) => kinds.get(f) === "date" && dateParts(cache, f) !== null).map(({ f }) => f)
  const name = (f: number) => cache.fields[f].name
  const out: PivotRecipe[] = []
  const add = (r: PivotRecipe) => {
    if (out.length < limit && !out.some((o) => o.title === r.title)) out.push(r)
  }
  const sumOf = (f: number) => `Sum of ${name(f)}`
  // One category by each number (the first number with every category first, then the others).
  for (const n of numbers.slice(0, 1))
    for (const c of categories) add({ title: `${sumOf(n)} by ${name(c)}`, rows: [c], cols: [], values: [n] })
  for (const d of dates)
    for (const n of numbers.slice(0, 2))
      add({
        title: `${sumOf(n)} by ${name(d)}`,
        rows: [d],
        cols: [],
        values: [n],
        dates: { field: d, parts: dateParts(cache, d)! },
      })
  for (const n of numbers.slice(1))
    for (const c of categories.slice(0, 3))
      add({ title: `${sumOf(n)} by ${name(c)}`, rows: [c], cols: [], values: [n] })
  // Two numbers side by side, by the first category.
  if (numbers.length > 1 && categories.length)
    add({
      title: `${sumOf(numbers[0])} and ${sumOf(numbers[1])} by ${name(categories[0])}`,
      rows: [categories[0]],
      cols: [],
      values: numbers.slice(0, 2),
    })
  // Two categories crossed: the smaller one across the columns.
  if (categories.length > 1 && numbers.length) {
    const [a, b] = categories
    if (items(a) <= 12)
      add({ title: `${sumOf(numbers[0])} by ${name(b)} and ${name(a)}`, rows: [b], cols: [a], values: [numbers[0]] })
  }
  // No numbers: how many rows each item has.
  if (!numbers.length)
    for (const c of categories) {
      add({ title: `Count of ${name(c)} by ${name(c)}`, rows: [c], cols: [], values: [c] })
      const other = categories.find((x) => x !== c)
      if (other !== undefined && items(c) <= 12)
        add({
          title: `Count of ${name(other)} by ${name(other)} and ${name(c)}`,
          rows: [other],
          cols: [c],
          values: [other],
        })
    }
  return out
}

/** How a date field is grouped for a suggestion: by months, with years when it spans more than one. */
function dateParts(cache: PivotCache, f: number): DatePart[] | null {
  const serials = fieldShared(cache, f).filter((v): v is number => typeof v === "number")
  if (serials.length < 2) return null
  const years = new Set(serials.map((s) => serialToDate(s).year))
  const months = new Set(serials.map((s) => serialToDate(s).year * 12 + serialToDate(s).month))
  if (months.size < 2) return null
  return years.size > 1 ? ["months", "years"] : ["months"]
}

/** The PivotTable a suggestion makes (at `at`), and its cache with the dates grouped as the suggestion groups them. */
export function pivotFromRecipe(
  name: string,
  cache: PivotCache,
  at: CellAddress,
  recipe: PivotRecipe,
): { table: PivotTable; cache: PivotCache } {
  return withRecipe(newPivotTable(name, cache, at), cache, recipe)
}

/**
 * PivotTable `start` laid out as the recipe says, from scratch (as after Clear All: its name, place, style and options
 * stay), and its cache with the dates grouped as the recipe groups them (`cache` itself when nothing is grouped).
 */
export function withRecipe(
  start: PivotTable,
  cache: PivotCache,
  recipe: PivotRecipe,
): { table: PivotTable; cache: PivotCache } {
  let grouped = cache
  let rows = recipe.rows
  let cols = recipe.cols
  /** Where a field of the recipe ended up: a grouped date field becomes its parts, the largest outermost. */
  let placedOf = (f: number) => [f]
  const dates = recipe.dates
  if (dates?.parts.length) {
    const next = withRangeGroup(cache, dates.field, { start: null, end: null, by: 1, parts: dates.parts })
    if (typeof next !== "string") {
      grouped = next
      const parts = derivedDateGroups(next, dates.field).map((d) => d.index)
      // The largest part outermost; one part grouped alone stays on the field itself.
      const placed = parts.length ? [...parts].reverse() : [dates.field]
      placedOf = (f) => (f === dates.field ? placed : [f])
      rows = rows.flatMap(placedOf)
      cols = cols.flatMap(placedOf)
    }
  }
  let table = withAllCleared(start, grouped)
  for (const f of recipe.pages ?? []) table = withFieldIn(table, grouped, f, "filters")
  for (const f of rows) table = withFieldIn(table, grouped, f, "rows")
  for (const f of cols) table = withFieldIn(table, grouped, f, "cols")
  recipe.values.forEach((f, i) => {
    table = withFieldIn(table, grouped, f, "values")
    const fn = recipe.fns?.[i]
    const index = table.values.length - 1
    if (fn && table.values[index]?.fn !== fn) {
      const next = withValueSettings(table, grouped, index, { name: "", fn })
      if (typeof next !== "string") table = next
    }
  })
  // Sort and Top N on the innermost field a date field became (where its items are listed).
  const target = (f: number) => {
    const placed = placedOf(f)
    const g = placed[placed.length - 1]
    return rows.includes(g) || cols.includes(g) ? g : null
  }
  const sort = recipe.sort
  const sortField = sort ? target(sort.field) : null
  if (sort && sortField !== null) {
    if (sort.by === "label") table = withSort(table, grouped, sortField, sort.order)
    else if (table.values[sort.by]) table = withValueSort(table, sortField, sort.by, sort.order)
  }
  const top = recipe.top
  const topField = top ? target(top.field) : null
  if (top && topField !== null && table.values[top.measure] && top.count > 0)
    table = withFieldFilter(table, topField, {
      kind: "top",
      by: "count",
      top: !top.bottom,
      measure: top.measure,
      value: Math.floor(top.count),
    })
  return { table, cache: grouped }
}
