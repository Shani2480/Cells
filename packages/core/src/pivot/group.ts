// PivotTable grouping (Analyze › Group Selection / Group Field / Ungroup), as Excel 365 does it:
// - numbers are grouped in ranges on the field itself ("1-5", "6-10", with "<1" and ">21" for values outside);
// - dates are grouped by Seconds … Years into fields of their own, "Months (Date)", the base field leaving the axis;
// - other items are gathered by Group Selection into a field of its own, "Product2", with Group1, Group2 ….
// The labels and ranges are Excel's, read from io's fixtures/excel/pivot-group.xlsx and from Excel's own groupItems.
import { serialToDate, serialToTime } from "../formula/dates"
import type { PivotCache, PivotCacheField, PivotFieldGroup, PivotGroupBy, PivotValue } from "./model"

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

/** The field holds a value in each record (not a calculated field or a group field of its own). */
export function inRecords(field: PivotCacheField | undefined, f: number): boolean {
  return !!field && field.formula === undefined && (!field.group || field.group.base === f)
}

const sharedCache = new WeakMap<PivotCache, Map<number, PivotValue[]>>()
const recordCache = new WeakMap<PivotCache, Map<number, Int32Array>>()

/**
 * The items a field's item indexes (`x`) point at: its group items when grouped, else the cache's shared items, or the
 * column's distinct values in the order first met when it keeps none.
 */
export function fieldShared(cache: PivotCache, f: number): PivotValue[] {
  const field = cache.fields[f]
  if (!field) return []
  if (field.group) return field.group.items
  if (field.items) return field.items
  let map = sharedCache.get(cache)
  if (!map) sharedCache.set(cache, (map = new Map()))
  let out = map.get(f)
  if (!out) {
    out = []
    const seen = new Set<string>()
    if (inRecords(field, f))
      for (const rec of cache.records) {
        const v = rec[f] ?? null
        const k = keyOf(v)
        if (!seen.has(k)) {
          seen.add(k)
          out.push(v)
        }
      }
    map.set(f, out)
  }
  return out
}

/** The item index (into `fieldShared`) of each record for field `f`; -1 where it has none. */
export function recordXs(cache: PivotCache, f: number): Int32Array {
  let map = recordCache.get(cache)
  if (!map) recordCache.set(cache, (map = new Map()))
  let out = map.get(f)
  if (out) return out
  const field = cache.fields[f]
  out = new Int32Array(cache.records.length).fill(-1)
  const group = field?.group
  if (group?.kind === "range") {
    const at = rangeIndex(group)
    cache.records.forEach((rec, i) => (out![i] = at(rec[group.base] ?? null)))
  } else if (group?.kind === "discrete") {
    const base = recordXs(cache, group.base)
    for (let i = 0; i < base.length; i++) out[i] = base[i] >= 0 ? (group.map[base[i]] ?? -1) : -1
  } else if (field && inRecords(field, f)) {
    const index = new Map<string, number>()
    fieldShared(cache, f).forEach((v, x) => index.has(keyOf(v)) || index.set(keyOf(v), x))
    cache.records.forEach((rec, i) => (out![i] = index.get(keyOf(rec[f] ?? null)) ?? -1))
  }
  map.set(f, out)
  return out
}

/** Day of the year in a leap year's calendar (1 Jan 0 … 31 Dec 365): Excel's Days items. */
function leapDayIndex(month: number, day: number): number {
  const before = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335]
  return before[month - 1] + day - 1
}

/** Where a value falls among a range group's items (0 "<start", the last ">end"). */
function rangeIndex(g: Extract<PivotFieldGroup, { kind: "range" }>): (v: PivotValue) => number {
  const last = g.items.length - 1
  return (v) => {
    if (typeof v !== "number") return 0
    if (v < g.start) return 0
    if (v > g.end) return last
    const inner = last - 1
    switch (g.by) {
      case "range":
      case "days":
        if (g.by === "days" && g.interval <= 1) {
          const d = serialToDate(Math.floor(v))
          return Math.min(leapDayIndex(d.month, d.day) + 1, inner)
        }
        return (
          Math.min(
            Math.floor((g.by === "days" ? Math.floor(v) - g.start : v - g.start) / g.interval + 1e-9),
            inner - 1,
          ) + 1
        )
      case "seconds":
        return serialToTime(v).second + 1
      case "minutes":
        return serialToTime(v).minute + 1
      case "hours":
        return serialToTime(v).hour + 1
      case "months":
        return serialToDate(Math.floor(v)).month
      case "quarters":
        return Math.ceil(serialToDate(Math.floor(v)).month / 3)
      case "years":
        return Math.min(serialToDate(Math.floor(v)).year - serialToDate(Math.floor(g.start)).year + 1, inner)
    }
  }
}

// ── Labels ──

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

/** A number as a group label shows it (General: no trailing zeros, at most 15 digits). */
const num = (x: number) => String(Number(x.toPrecision(15)))

/** A date as the group labels write it: m/d/yyyy. */
function mdy(serial: number): string {
  const d = serialToDate(Math.floor(serial))
  return `${d.month}/${d.day}/${d.year}`
}

const pad = (n: number) => (n < 10 ? `:0${n}` : `:${n}`)

/**
 * The group items Excel makes for a range group: "<start", each range or date part, ">end". `wholeNumbers`: the data,
 * start and interval are whole numbers, so ranges read 1-5, 6-10 (else 0-100, 100-200). `blank`: some record is empty,
 * and goes with the first item ("<0 or (blank)").
 */
export function rangeGroupItems(
  by: PivotGroupBy,
  start: number,
  end: number,
  interval: number,
  wholeNumbers: boolean,
  blank: boolean,
): string[] {
  const dates = by !== "range"
  const first = `<${dates ? mdy(start) : num(start)}${blank ? " or (blank)" : ""}`
  const middle: string[] = []
  let after = dates ? mdy(end) : ""
  switch (by) {
    case "range": {
      const n = Math.max(1, Math.ceil((end - start) / interval - 1e-9))
      for (let k = 0; k < n; k++) {
        const lo = start + k * interval
        let hi = wholeNumbers ? lo + interval - 1 : lo + interval
        if (k === n - 1) hi = Math.max(hi, end)
        middle.push(`${num(lo)}-${num(hi)}`)
      }
      after = num(start + n * interval)
      break
    }
    case "seconds":
    case "minutes":
      for (let k = 0; k < 60; k++) middle.push(pad(k))
      break
    case "hours":
      for (let k = 0; k < 24; k++) middle.push(`${k % 12 === 0 ? 12 : k % 12} ${k < 12 ? "AM" : "PM"}`)
      break
    case "days":
      if (interval <= 1) {
        for (let m = 1; m <= 12; m++) {
          const days = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
          for (let d = 1; d <= days; d++) middle.push(`${d}-${MONTHS[m - 1]}`)
        }
      } else {
        const n = Math.max(1, Math.ceil((end - start) / interval - 1e-9))
        for (let k = 0; k < n; k++) {
          const lo = start + k * interval
          middle.push(`${mdy(lo)} - ${mdy(lo + interval - 1)}`)
        }
      }
      break
    case "months":
      middle.push(...MONTHS)
      break
    case "quarters":
      middle.push("Qtr1", "Qtr2", "Qtr3", "Qtr4")
      break
    case "years": {
      const y0 = serialToDate(Math.floor(start)).year
      const y1 = serialToDate(Math.floor(end)).year
      for (let y = y0; y <= y1; y++) middle.push(String(y))
      break
    }
  }
  return [first, ...middle, `>${after}`]
}

/** The words Excel's Grouping dialog and field names use for each date part. */
export const DATE_GROUP_NAMES: Record<Exclude<PivotGroupBy, "range">, string> = {
  seconds: "Seconds",
  minutes: "Minutes",
  hours: "Hours",
  days: "Days",
  months: "Months",
  quarters: "Quarters",
  years: "Years",
}

/** The date parts from the smallest: Excel makes their fields in this order, and puts the largest outermost. */
export const DATE_PARTS: Exclude<PivotGroupBy, "range">[] = [
  "seconds",
  "minutes",
  "hours",
  "days",
  "months",
  "quarters",
  "years",
]

// ── What a field can be grouped by ──

export type GroupKind = "numbers" | "dates" | "items"

/** The numbers (or date serials) of a field's records, and whether some record is empty. */
function numbersOf(cache: PivotCache, f: number): { values: number[]; text: boolean; blank: boolean } {
  const values: number[] = []
  let text = false
  let blank = false
  for (const rec of cache.records) {
    const v = rec[f] ?? null
    if (v === null || v === "") blank = true
    else if (typeof v === "number") values.push(v)
    else text = true
  }
  return { values, text, blank }
}

/** How a field groups: dates (the column holds dates), numbers (only numbers), else items by Group Selection. */
export function groupKind(cache: PivotCache, f: number): GroupKind {
  const field = cache.fields[f]
  if (!field || !inRecords(field, f) || (field.group && field.group.kind === "discrete")) return "items"
  const { values, text } = numbersOf(cache, f)
  if (text || !values.length) return "items"
  return field.dates ? "dates" : "numbers"
}

export type GroupSettings = {
  /** Null: Auto (the smallest value; for dates, its day). */
  start: number | null
  /** Null: Auto (the largest value; for dates, the day after it). */
  end: number | null
  /** Numbers: the size of each range. Dates: days per item, when Days is the only part. */
  by: number
  /** Dates: the parts to group by. */
  parts?: Exclude<PivotGroupBy, "range">[]
}

/** The Grouping dialog's values for a field: its group's, or Excel's defaults (Auto, By 10, Months). */
export function groupSettings(cache: PivotCache, f: number): GroupSettings & { autoStart: number; autoEnd: number } {
  const kind = groupKind(cache, f)
  const { values } = numbersOf(cache, f)
  let lo = Infinity
  let hi = -Infinity
  for (const v of values) {
    if (v < lo) lo = v
    if (v > hi) hi = v
  }
  const autoStart = kind === "dates" ? Math.floor(lo) : lo
  const autoEnd = kind === "dates" ? Math.floor(hi) + 1 : hi
  const groups = derivedDateGroups(cache, f)
  const own = cache.fields[f]?.group
  const g = kind === "dates" ? groups[0]?.group : own?.kind === "range" ? own : undefined
  if (!g || g.kind !== "range")
    return { start: null, end: null, by: kind === "dates" ? 1 : 10, parts: ["months"], autoStart, autoEnd }
  return {
    start: g.autoStart ? null : g.start,
    end: g.autoEnd ? null : g.end,
    by: g.interval,
    parts:
      kind === "dates"
        ? groups.map((d) => (d.group as { by: PivotGroupBy }).by as Exclude<PivotGroupBy, "range">)
        : undefined,
    autoStart,
    autoEnd,
  }
}

/** The fields Excel made from date field `f` (Months (Date) …), smallest part first. */
export function derivedDateGroups(cache: PivotCache, f: number): { index: number; group: PivotFieldGroup }[] {
  const out: { index: number; group: PivotFieldGroup }[] = []
  cache.fields.forEach((field, i) => {
    if (i !== f && field.group?.kind === "range" && field.group.base === f) out.push({ index: i, group: field.group })
  })
  return out
}

export const GROUP_ERROR = "Cannot group that selection."
export const GROUP_RANGE_ERROR = "The ending value must be greater than the starting value."

/**
 * The cache with field `f` grouped as the Grouping dialog says: numbers in ranges on the field itself; dates into a
 * field per part (any made before are replaced). Excel's message when it can't.
 */
export function withRangeGroup(cache: PivotCache, f: number, settings: GroupSettings): PivotCache | string {
  const kind = groupKind(cache, f)
  if (kind === "items") return GROUP_ERROR
  const defaults = groupSettings(cache, f)
  const start = settings.start ?? defaults.autoStart
  const end = settings.end ?? defaults.autoEnd
  if (!(end > start)) return GROUP_RANGE_ERROR
  const { values, blank } = numbersOf(cache, f)
  const base = {
    kind: "range" as const,
    base: f,
    start,
    end,
    autoStart: settings.start === null,
    autoEnd: settings.end === null,
  }
  if (kind === "numbers") {
    if (!(settings.by > 0)) return GROUP_ERROR
    const whole = Number.isInteger(start) && Number.isInteger(settings.by) && values.every((v) => Number.isInteger(v))
    const items = rangeGroupItems("range", start, end, settings.by, whole, blank)
    const fields = [...cache.fields]
    fields[f] = { ...fields[f], group: { ...base, by: "range", interval: settings.by, items } }
    return { ...cache, fields }
  }
  const parts = DATE_PARTS.filter((p) => settings.parts?.includes(p))
  if (!parts.length) return GROUP_ERROR
  const onlyDays = parts.length === 1 && parts[0] === "days"
  const interval = onlyDays ? Math.max(1, Math.round(settings.by)) : 1
  let next = withoutDateGroups(cache, f)
  const name = cache.fields[f].name
  const added: PivotCacheField[] = parts.map((by) => ({
    name: `${DATE_GROUP_NAMES[by]} (${name})`,
    items: null,
    group: { ...base, by, interval, items: rangeGroupItems(by, start, end, interval, false, blank) },
  }))
  next = { ...next, fields: [...next.fields, ...added] }
  return next
}

/** The cache without the fields made from date field `f`. */
export function withoutDateGroups(cache: PivotCache, f: number): PivotCache {
  const drop = new Set(derivedDateGroups(cache, f).map((d) => d.index))
  return drop.size ? withoutFields(cache, drop) : cache
}

/** The cache without some fields of no records (group fields), other groups' bases moved down to match. */
function withoutFields(cache: PivotCache, drop: Set<number>): PivotCache {
  const moved = (i: number) => i - [...drop].filter((d) => d < i).length
  const fields = cache.fields.flatMap((field, i) => {
    if (drop.has(i)) return []
    if (!field.group) return [field]
    return [{ ...field, group: { ...field.group, base: moved(field.group.base) } }]
  })
  return { ...cache, fields }
}

type Discrete = Extract<PivotFieldGroup, { kind: "discrete" }>

/**
 * The levels of items gathered from field `root` by Group Selection, in order: Product2, then Product3 made from
 * Product2's items. Excel bases every level on the source field itself, each mapping its items.
 */
export function discreteLevels(cache: PivotCache, root: number): number[] {
  const out: number[] = []
  cache.fields.forEach((k, i) => {
    if (i !== root && k.group?.kind === "discrete" && k.group.base === root) out.push(i)
  })
  return out
}

/** For field `f` (the source or one of its levels): its source field, its level (0 the source) and the levels. */
function levelOf(cache: PivotCache, f: number): { root: number; level: number; levels: number[] } {
  const g = cache.fields[f]?.group
  const root = g?.kind === "discrete" ? g.base : f
  const levels = discreteLevels(cache, root)
  return { root, level: levels.indexOf(f) + 1, levels }
}

/** The item of level `level` each of the source's items falls in (level 0: the item itself). */
function itemsAt(cache: PivotCache, root: number, levels: number[], level: number): number[] {
  const count = fieldShared(cache, root).length
  if (level === 0) return Array.from({ length: count }, (_, x) => x)
  const g = cache.fields[levels[level - 1]].group as Discrete
  return Array.from({ length: count }, (_, x) => g.map[x] ?? -1)
}

/**
 * Group Selection on the items `xs` of field `f` (indexes into `fieldShared`; at least two, as Excel asks): they
 * become one group, GroupN (numbered across the levels), in the level above `f` — made the first time as "Product2",
 * "Product3", the items not picked keeping their own. Returns the cache and that level's field.
 */
export function withItemsGrouped(
  cache: PivotCache,
  f: number,
  xs: readonly number[],
): { cache: PivotCache; field: number } | string {
  const picked = new Set(xs)
  if (picked.size < 2 || !cache.fields[f]) return GROUP_ERROR
  const { root, level, levels } = levelOf(cache, f)
  const fields = [...cache.fields]
  // The source keeps its shared items: the levels point at them.
  const rootItems = fieldShared(cache, root)
  if (inRecords(fields[root], root) && !fields[root].items && !fields[root].group)
    fields[root] = { ...fields[root], items: [...rootItems] }
  const below = itemsAt(cache, root, levels, level)
  const labels = fieldShared(cache, f)
  const taken = new Set(
    levels.flatMap((l) => (cache.fields[l].group as Discrete).items.map((v) => String(v).toLowerCase())),
  )
  let n = 1
  while (taken.has(`group${n}`)) n++
  const groupName = `Group${n}`
  const upper = levels[level]
  if (upper === undefined) {
    const items: PivotValue[] = []
    const own = new Map<number, number>()
    labels.forEach((v, x) => {
      if (picked.has(x)) return
      own.set(x, items.length)
      items.push(v)
    })
    const g = items.length
    items.push(groupName)
    const map = below.map((x) => (picked.has(x) ? g : (own.get(x) ?? -1)))
    const used = new Set(fields.map((k) => k.name.toLowerCase()))
    let k = 2
    while (used.has(`${fields[root].name}${k}`.toLowerCase())) k++
    fields.push({
      name: `${fields[root].name}${k}`,
      items: null,
      group: tidyDiscrete({ kind: "discrete", base: root, map, items }),
    })
    return { cache: { ...cache, fields }, field: fields.length - 1 }
  }
  const old = fields[upper].group as Discrete
  const items = [...old.items, groupName]
  const map = below.map((x, rx) => (picked.has(x) ? items.length - 1 : (old.map[rx] ?? -1)))
  fields[upper] = { ...fields[upper], group: tidyDiscrete({ ...old, map, items }) }
  return { cache: { ...cache, fields }, field: upper }
}

/** A discrete group without the items no base item points at any more. */
function tidyDiscrete(g: Discrete): Discrete {
  const used = new Set(g.map)
  const newIndex: number[] = []
  const items: PivotValue[] = []
  g.items.forEach((v, i) => {
    if (!used.has(i)) return
    newIndex[i] = items.length
    items.push(v)
  })
  return { ...g, items, map: g.map.map((i) => newIndex[i] ?? -1) }
}

/**
 * Ungroup on field `f`: number ranges go; a date field's part fields go (from any of them); in a field made by Group
 * Selection the groups `xs` (all when none) split into their items again, and the field goes when none is left.
 * Returns the cache and the field the PivotTables show instead (the base), or null when there's nothing to ungroup.
 */
export function withUngrouped(
  cache: PivotCache,
  f: number,
  xs?: readonly number[],
): { cache: PivotCache; base: number; removed: boolean } | null {
  const field = cache.fields[f]
  const g = field?.group
  if (!field) return null
  if (!g) {
    const derived = derivedDateGroups(cache, f)
    return derived.length ? { cache: withoutDateGroups(cache, f), base: f, removed: false } : null
  }
  if (g.kind === "range") {
    if (g.base === f) {
      const fields = [...cache.fields]
      const { group: _gone, ...rest } = fields[f]
      fields[f] = rest
      return { cache: { ...cache, fields }, base: f, removed: false }
    }
    return { cache: withoutDateGroups(cache, g.base), base: g.base, removed: true }
  }
  // A level made by Group Selection: its groups split into the items of the level below.
  const { root, level, levels } = levelOf(cache, f)
  const below = itemsAt(cache, root, levels, level - 1)
  const belowLabels = fieldShared(cache, level === 1 ? root : levels[level - 2])
  const split = new Set(xs?.length ? xs : g.items.map((_, i) => i))
  const items = [...g.items]
  const own = new Map<number, number>()
  const map = g.map.map((gi, rx) => {
    if (!split.has(gi)) return gi
    const x = below[rx]
    if (!own.has(x)) {
      own.set(x, items.length)
      items.push(belowLabels[x] ?? null)
    }
    return own.get(x)!
  })
  const next = tidyDiscrete({ ...g, items, map })
  // Nothing grouped any more: every item is one item of the level below, under its own label.
  const plain = next.items.every((v, i) => {
    const of = new Set(next.map.flatMap((gi, rx) => (gi === i ? [below[rx]] : [])))
    return of.size === 1 && keyOf(v) === keyOf(belowLabels[[...of][0]] ?? null)
  })
  const back = level === 1 ? root : levels[level - 2]
  if (plain) return { cache: withoutFields(cache, new Set([f])), base: back, removed: true }
  const fields = [...cache.fields]
  fields[f] = { ...field, group: next }
  return { cache: { ...cache, fields }, base: back, removed: false }
}

/** The group a field shows when items are added to its records after a refresh: new base items get their own. */
export function refreshedDiscrete(
  g: Extract<PivotFieldGroup, { kind: "discrete" }>,
  oldBase: PivotValue[],
  newBase: PivotValue[],
  base: number,
): PivotFieldGroup {
  const oldX = new Map(oldBase.map((v, x) => [keyOf(v), x]))
  const items = [...g.items]
  const map = newBase.map((v) => {
    const x = oldX.get(keyOf(v))
    if (x !== undefined && g.map[x] !== undefined && g.map[x] >= 0) return g.map[x]
    items.push(v)
    return items.length - 1
  })
  return tidyDiscrete({ ...g, base, items, map })
}
