import {
  displayText,
  fieldShared,
  groupKind,
  inRecords,
  PIVOT_FN_LABEL,
  PIVOT_FNS,
  type DatePart,
  type PivotCache,
  type PivotFn,
  type PivotRecipe,
} from "@cells/core"
import type { Request } from "./prompts"

/**
 * "Describe the PivotTable you want" (docs/ai.md, docs/pivot-plan.md step 8): the request and how the reply is read.
 * The model only proposes a layout — which columns go on the rows, columns, values and filters, how values are
 * summarized, dates grouped, one sort and one Top N — and Cells' own PivotTable engine works out every number.
 * It is sent the columns' names, what kind of data each holds and a few of their values, never whole rows.
 */

/** How many different values of a column the request shows. */
const SAMPLES = 6

/** The columns a description can use: the data's own, not groups or calculated fields made in a PivotTable. */
export function describableFields(cache: PivotCache): number[] {
  return cache.fields.map((_, i) => i).filter((i) => inRecords(cache.fields[i], i))
}

/** One line per column: "- Region: text, 3 different values (North, South, East)". */
export function describeColumns(cache: PivotCache): string {
  return describableFields(cache)
    .map((f) => {
      const field = cache.fields[f]
      const values = fieldShared(cache, f).filter((v) => v !== null && v !== "")
      const kind = groupKind(cache, f)
      if (!values.length) return `- ${field.name}: empty`
      if (kind === "dates" || kind === "numbers") {
        const nums = values.filter((v): v is number => typeof v === "number")
        const lo = Math.min(...nums)
        const hi = Math.max(...nums)
        const show = (n: number) => (kind === "dates" ? displayText(n, "yyyy-mm-dd") : displayText(n, "General"))
        return `- ${field.name}: ${kind === "dates" ? "dates" : "numbers"} from ${show(lo)} to ${show(hi)}`
      }
      const some = values.slice(0, SAMPLES).map((v) => displayText(v ?? undefined, "General"))
      const more = values.length > SAMPLES ? ", …" : ""
      return `- ${field.name}: text, ${values.length} different value${values.length === 1 ? "" : "s"} (${some.join(", ")}${more})`
    })
    .join("\n")
}

const FN_WORDS = PIVOT_FNS.join(", ")
const DATE_WORDS: DatePart[] = ["days", "months", "quarters", "years"]

/** The request: a PivotTable layout for the description, as JSON. */
export function pivotRequest(cache: PivotCache, description: string, rows: number): Request {
  return {
    system: `You work inside Cells, a spreadsheet that follows Microsoft Excel. Turn the person's description into an
Excel PivotTable layout over their data. Use only the column names listed, written exactly as listed. Reply with JSON
only, no code fence:
{"title": "a short title, e.g. Sum of Amount by Region",
 "rows": ["column", ...], "columns": ["column", ...], "filters": ["column", ...],
 "values": [{"field": "column", "summary": "<one of: ${FN_WORDS}>"}],
 "group_dates": {"field": "a date column on rows or columns", "by": ["<some of: ${DATE_WORDS.join(", ")}>"]},
 "sort": {"field": "a column on rows or columns", "by": "label" or "<a column in values>", "order": "ascending" or "descending"},
 "top": {"field": "a column on rows or columns", "count": 10, "bottom": false, "by": "<a column in values>"}}
Leave out what the description doesn't ask for (group_dates, sort, top are optional). Numbers are usually summed,
text counted. "Biggest first" means sort by the value, descending. "Per month/quarter/year" on a date column means
group_dates. If the description can't be done with this data, reply {"problem": "one sentence on why"}.`,
    prompt: description.trim(),
    context: `The data has ${rows} rows and these columns:\n${describeColumns(cache)}`,
  }
}

export type PivotReply = { recipe: PivotRecipe } | { problem: string }

/** The first JSON object in a reply (models sometimes add words around it). */
function jsonIn(reply: string): Record<string, unknown> | null {
  const start = reply.indexOf("{")
  const end = reply.lastIndexOf("}")
  if (start < 0 || end <= start) return null
  try {
    const value = JSON.parse(reply.slice(start, end + 1)) as unknown
    return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null
  } catch {
    return null
  }
}

const str = (v: unknown) => (typeof v === "string" ? v.trim() : "")
const list = (v: unknown): unknown[] => (Array.isArray(v) ? v : v === undefined || v === null || v === "" ? [] : [v])
const obj = (v: unknown): Record<string, unknown> | null =>
  v && typeof v === "object" && !Array.isArray(v) ? (v as Record<string, unknown>) : null

/** Reads the layout the model proposed, checking every column against the data. */
export function readPivotReply(reply: string, cache: PivotCache): PivotReply {
  const json = jsonIn(reply)
  if (!json) return { problem: "The AI didn't answer with a PivotTable. Try describing it another way." }
  if (str(json.problem)) return { problem: str(json.problem) }
  const usable = describableFields(cache)
  const byName = new Map(usable.map((f) => [cache.fields[f].name.trim().toLowerCase(), f]))
  const unknown: string[] = []
  const field = (v: unknown): number | null => {
    const name = str(v) || str(obj(v)?.field)
    if (!name) return null
    const f = byName.get(name.toLowerCase())
    if (f === undefined) {
      if (!unknown.includes(name)) unknown.push(name)
      return null
    }
    return f
  }
  const fields = (v: unknown) =>
    list(v)
      .map(field)
      .filter((f): f is number => f !== null)
  // A column is on one of rows, columns and filters (the first named wins), as in a PivotTable.
  const placed = new Set<number>()
  const take = (fs: number[]) => fs.filter((f) => !placed.has(f) && (placed.add(f), true))
  const rows = take(fields(json.rows))
  const cols = take(fields(json.columns ?? json.cols))
  const pages = take(fields(json.filters))
  const values: number[] = []
  const fns: (PivotFn | undefined)[] = []
  for (const v of list(json.values)) {
    const f = field(v)
    if (f === null) continue
    const said = str(obj(v)?.summary).toLowerCase()
    const fn = PIVOT_FNS.find((x) => x.toLowerCase() === said || PIVOT_FN_LABEL[x].toLowerCase() === said)
    values.push(f)
    fns.push(fn)
  }
  if (unknown.length)
    return {
      problem: `The AI used ${unknown.length === 1 ? "a column" : "columns"} the data doesn't have: ${unknown.join(", ")}. Try again, or name the columns as they are written.`,
    }
  if (!rows.length && !cols.length && !values.length)
    return { problem: "The AI didn't put any columns in the PivotTable. Try describing it another way." }
  const recipe: PivotRecipe = { title: str(json.title), rows, cols, values }
  if (fns.some(Boolean)) recipe.fns = fns
  if (pages.length) recipe.pages = pages
  const onAxis = (f: number | null): f is number => f !== null && (rows.includes(f) || cols.includes(f))
  // Dates grouped: only a date column on the rows or columns.
  const dates = obj(json.group_dates ?? json.dates)
  const dateField = dates ? field(dates.field) : null
  if (dates && onAxis(dateField) && groupKind(cache, dateField) === "dates") {
    const parts = DATE_WORDS.filter((p) => list(dates.by).some((b) => str(b).toLowerCase() === p))
    if (parts.length) recipe.dates = { field: dateField, parts }
  }
  /** A value named by its column (the first value made from it), or by its title ("Sum of Amount"). */
  const valueIndex = (v: unknown): number => {
    const name = str(v).toLowerCase()
    if (!name) return values.length ? 0 : -1
    const byField = values.findIndex((f) => cache.fields[f].name.toLowerCase() === name)
    if (byField >= 0) return byField
    return values.findIndex(
      (f, i) => `${PIVOT_FN_LABEL[fns[i] ?? "sum"]} of ${cache.fields[f].name}`.toLowerCase() === name,
    )
  }
  const sort = obj(json.sort)
  const sortField = sort ? field(sort.field) : null
  if (sort && onAxis(sortField)) {
    const order = str(sort.order).toLowerCase().startsWith("desc") ? "descending" : "ascending"
    const by = str(sort.by).toLowerCase()
    if (!by || by === "label" || by === "labels" || by === "name")
      recipe.sort = { field: sortField, by: "label", order }
    else {
      const index = valueIndex(sort.by)
      if (index >= 0) recipe.sort = { field: sortField, by: index, order }
    }
  }
  const top = obj(json.top)
  const topField = top ? field(top.field) : null
  const count = Number(top?.count)
  if (top && onAxis(topField) && Number.isFinite(count) && count >= 1) {
    const measure = valueIndex(top.by)
    if (measure >= 0) recipe.top = { field: topField, count: Math.floor(count), bottom: top.bottom === true, measure }
  }
  if (!recipe.title) recipe.title = recipeTitle(cache, recipe)
  return { recipe }
}

/** A title as Recommended PivotTables words them: "Sum of Amount by Region and Channel". */
export function recipeTitle(cache: PivotCache, r: PivotRecipe): string {
  const name = (f: number) => cache.fields[f].name
  const vals = r.values.map(
    (f, i) => `${PIVOT_FN_LABEL[r.fns?.[i] ?? (numeric(cache, f) ? "sum" : "count")]} of ${name(f)}`,
  )
  const by = [...r.rows, ...r.cols].map(name)
  const join = (xs: string[]) =>
    xs.length < 2 ? xs.join("") : `${xs.slice(0, -1).join(", ")} and ${xs[xs.length - 1]}`
  return [join(vals) || "PivotTable", by.length ? `by ${join(by)}` : ""].filter(Boolean).join(" ")
}

const numeric = (cache: PivotCache, f: number) =>
  fieldShared(cache, f).every((v) => v === null || v === "" || typeof v === "number")
