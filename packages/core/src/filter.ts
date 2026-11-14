import type { Range } from "./address"
import type { CellData } from "./cell-store"
import { serialToDate, serialWeekday } from "./formula/dates"
import { displayText } from "./format"
import { isDateFormat } from "./numfmt"
import { shiftRange } from "./conditional"
import type { AxisEntry } from "./ops"
import type { Sheet } from "./sheet"
import { compareSortText } from "./sort"
import type { StyleTable } from "./styles"
import { parseInput } from "./values"

/**
 * AutoFilter, as Excel filters (checked against Excel 365 on a column of mixed values):
 * - "equals" and the value list compare the text a cell shows (5.0000000001 shows as 5 and equals "5"), ignoring
 *   case and surrounding spaces, with * ? ~ wildcards; "=" alone is blank cells, "<>" alone the filled ones;
 * - "does not equal" a number hides only that number; otherwise it is the opposite of "equals";
 * - greater/less than a number or date look at numbers only (not text that reads as one); than a text, at text
 *   only (not TRUE/FALSE or errors);
 * - Top 10 and averages look at numbers; Top 10 keeps ties; a percentage rounds down (40 % of 8 numbers is 3).
 */
export type FilterOperator = "eq" | "ne" | "gt" | "ge" | "lt" | "le"
export type FilterCondition = { op: FilterOperator; value: string }

export type DateGroup = { year: number; month?: number; day?: number }

export type DynamicFilter =
  | "aboveAverage"
  | "belowAverage"
  | "today"
  | "yesterday"
  | "tomorrow"
  | "thisWeek"
  | "lastWeek"
  | "nextWeek"
  | "thisMonth"
  | "lastMonth"
  | "nextMonth"
  | "thisQuarter"
  | "lastQuarter"
  | "nextQuarter"
  | "thisYear"
  | "lastYear"
  | "nextYear"
  | "yearToDate"
  | `Q${1 | 2 | 3 | 4}`
  | `M${1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12}`

export type ColumnFilter =
  /** The checked items of the value list (shown texts), blanks, and checked dates by year, month or day. */
  | { kind: "values"; values: string[]; blanks?: boolean; dates?: DateGroup[] }
  /** Custom AutoFilter: one or two conditions, both (and) or either (or). */
  | { kind: "custom"; conditions: FilterCondition[]; and?: boolean }
  | { kind: "top10"; top: boolean; percent: boolean; count: number }
  | { kind: "dynamic"; type: DynamicFilter }
  /** Filter by Color: the cell's fill (cell) or font colour; null is No Fill / Automatic. */
  | { kind: "color"; cell: boolean; color: string | null }

/**
 * A sheet's AutoFilter: its range (the header row first), the filters by column offset in the range, and the last
 * sort made from a dropdown (its button shows the arrow).
 */
export type AutoFilter = {
  range: Range
  columns: Record<number, ColumnFilter>
  sort?: { col: number; descending: boolean }
}

const norm = (text: string) => text.trim().toLowerCase()

/** A filter pattern ("*an*", "a?c", "x~*y") as a regular expression over the whole (normalized) text. */
export function wildcardRegExp(pattern: string): RegExp {
  let out = ""
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i]
    if (c === "~" && i + 1 < pattern.length && "*?~".includes(pattern[i + 1])) {
      out += "\\" + pattern[++i]
    } else if (c === "*") out += "[\\s\\S]*"
    else if (c === "?") out += "[\\s\\S]"
    else out += c.replace(/[.+^${}()|[\]\\/-]/g, "\\$&")
  }
  return new RegExp(`^${out}$`, "i")
}

type Cell = { data: CellData | undefined; text: string; code?: string }

const isBlank = (data: CellData | undefined) => data?.value === undefined || data.value === ""

/** A criterion as a number, when it reads as one (dates and percentages too). */
function criterionNumber(text: string): number | null {
  const parsed = parseInput(text.trim())
  return parsed && typeof parsed.value === "number" && !parsed.formula ? parsed.value : null
}

function matchesCondition(cell: Cell, c: FilterCondition): boolean {
  const blank = isBlank(cell.data)
  const value = cell.data?.value
  if (c.value === "") return c.op === "eq" ? blank : c.op === "ne" ? !blank : false
  if (c.op === "eq" || c.op === "ne") {
    const number = criterionNumber(c.value)
    if (c.op === "ne" && number !== null) return !(typeof value === "number" && value === number)
    const hit = !blank && wildcardRegExp(norm(c.value)).test(norm(cell.text))
    return c.op === "eq" ? hit : !hit
  }
  if (blank) return false
  const number = criterionNumber(c.value)
  let order: number
  if (number !== null) {
    if (typeof value !== "number") return false
    order = value - number
  } else {
    if (typeof value !== "string") return false
    order = compareSortText(norm(value), norm(c.value))
  }
  switch (c.op) {
    case "gt":
      return order > 0
    case "ge":
      return order >= 0
    case "lt":
      return order < 0
    default:
      return order <= 0
  }
}

/** What a filter needs from the whole column before it can test cells (Top 10's cut-off, the average). */
type Prepared = { cut?: number; average?: number; today: number; values?: Set<string> }

function prepare(filter: ColumnFilter, cells: Cell[], today: number): Prepared {
  const numbers = () => cells.map((c) => c.data?.value).filter((v): v is number => typeof v === "number")
  if (filter.kind === "top10") {
    const list = numbers().sort((a, b) => (filter.top ? b - a : a - b))
    if (!list.length) return { today }
    const count = filter.percent ? Math.floor((list.length * filter.count) / 100) : Math.floor(filter.count)
    return { cut: list[Math.min(list.length, Math.max(1, count)) - 1], today }
  }
  if (filter.kind === "dynamic" && (filter.type === "aboveAverage" || filter.type === "belowAverage")) {
    const list = numbers()
    return { average: list.length ? list.reduce((a, b) => a + b, 0) / list.length : undefined, today }
  }
  if (filter.kind === "values") return { values: new Set(filter.values.map(norm)), today }
  return { today }
}

const quarterOf = (month: number) => Math.floor((month - 1) / 3)

function inDynamic(type: DynamicFilter, serial: number, today: number): boolean {
  const day = Math.floor(serial)
  const t = serialToDate(today)
  const d = serialToDate(day)
  // Weeks run Sunday to Saturday, as Excel's date filters.
  const weekStart = today - serialWeekday(today)
  const monthIndex = (x: { year: number; month: number }) => x.year * 12 + x.month - 1
  switch (type) {
    case "today":
      return day === today
    case "yesterday":
      return day === today - 1
    case "tomorrow":
      return day === today + 1
    case "thisWeek":
      return day >= weekStart && day < weekStart + 7
    case "lastWeek":
      return day >= weekStart - 7 && day < weekStart
    case "nextWeek":
      return day >= weekStart + 7 && day < weekStart + 14
    case "thisMonth":
      return monthIndex(d) === monthIndex(t)
    case "lastMonth":
      return monthIndex(d) === monthIndex(t) - 1
    case "nextMonth":
      return monthIndex(d) === monthIndex(t) + 1
    case "thisQuarter":
      return d.year * 4 + quarterOf(d.month) === t.year * 4 + quarterOf(t.month)
    case "lastQuarter":
      return d.year * 4 + quarterOf(d.month) === t.year * 4 + quarterOf(t.month) - 1
    case "nextQuarter":
      return d.year * 4 + quarterOf(d.month) === t.year * 4 + quarterOf(t.month) + 1
    case "thisYear":
      return d.year === t.year
    case "lastYear":
      return d.year === t.year - 1
    case "nextYear":
      return d.year === t.year + 1
    case "yearToDate":
      return d.year === t.year && day <= today
    default:
      if (type.startsWith("Q")) return quarterOf(d.month) === Number(type.slice(1)) - 1
      if (type.startsWith("M")) return d.month === Number(type.slice(1))
      return false
  }
}

function passes(cell: Cell, filter: ColumnFilter, p: Prepared, styles: StyleTable): boolean {
  const value = cell.data?.value
  switch (filter.kind) {
    case "values": {
      if (isBlank(cell.data)) return Boolean(filter.blanks)
      if (p.values!.has(norm(cell.text))) return true
      if (filter.dates?.length && typeof value === "number" && cell.code && isDateFormat(cell.code)) {
        const d = serialToDate(Math.floor(value))
        return filter.dates.some(
          (g) =>
            g.year === d.year &&
            (g.month === undefined || g.month === d.month) &&
            (g.day === undefined || g.day === d.day),
        )
      }
      return false
    }
    case "custom": {
      const results = filter.conditions.map((c) => matchesCondition(cell, c))
      return filter.and === false ? results.some(Boolean) : results.every(Boolean)
    }
    case "top10":
      return typeof value === "number" && p.cut !== undefined && (filter.top ? value >= p.cut : value <= p.cut)
    case "dynamic":
      if (filter.type === "aboveAverage" || filter.type === "belowAverage") {
        if (typeof value !== "number" || p.average === undefined) return false
        return filter.type === "aboveAverage" ? value > p.average : value < p.average
      }
      return typeof value === "number" && inDynamic(filter.type, value, p.today)
    case "color": {
      const style = styles.get(cell.data?.style ?? 0)
      const color = ((filter.cell ? style.fill : style.fontColor) ?? null)?.toUpperCase() ?? null
      return color === (filter.color?.toUpperCase() ?? null)
    }
  }
}

/** The cells of one column of the filter's data rows, with the text each shows. */
export function filterColumnCells(sheet: Sheet, styles: StyleTable, range: Range, col: number): Cell[] {
  const cells: Cell[] = []
  for (let r = range.r0 + 1; r <= range.r1; r++) {
    const data = sheet.cells.get(r, col)
    const code = styles.get(sheet.styleAt(r, col)).numberFormat
    cells.push({ data, code, text: displayText(data?.value, code) })
  }
  return cells
}

/** For each data row of the filter (from the row under the headers), whether it passes every column's filter. */
export function filterVisibility(sheet: Sheet, styles: StyleTable, filter: AutoFilter, today: number): boolean[] {
  const { range } = filter
  const shown = new Array<boolean>(Math.max(0, range.r1 - range.r0)).fill(true)
  for (const [offset, column] of Object.entries(filter.columns)) {
    const cells = filterColumnCells(sheet, styles, range, range.c0 + Number(offset))
    const p = prepare(column, cells, today)
    cells.forEach((cell, i) => {
      if (shown[i] && !passes(cell, column, p, styles)) shown[i] = false
    })
  }
  return shown
}

/** The row changes that show and hide the filter's data rows (only those whose state changes). */
export function filterRowEntries(sheet: Sheet, styles: StyleTable, filter: AutoFilter, today: number): AxisEntry[] {
  const shown = filterVisibility(sheet, styles, filter, today)
  const entries: AxisEntry[] = []
  shown.forEach((visible, i) => {
    const row = filter.range.r0 + 1 + i
    if (sheet.rows.isHidden(row) === !visible) return
    entries.push({
      index: row,
      size: sheet.rows.hasCustomSize(row) ? sheet.rows.storedSize(row) : null,
      hidden: !visible,
      ...(sheet.rows.isAuto(row) ? { auto: true as const } : {}),
    })
  })
  return entries
}

/** A deep copy (for undo snapshots). */
export function copyAutoFilter(filter: AutoFilter | null): AutoFilter | null {
  return filter ? (JSON.parse(JSON.stringify(filter)) as AutoFilter) : null
}

/**
 * The AutoFilter after rows or columns are inserted (count > 0) or deleted (count < 0) at `at`: its range follows as
 * a merge's does, the filters of deleted columns go and those to their right move with their columns. Null when its
 * whole range was deleted.
 */
export function shiftAutoFilter(
  filter: AutoFilter,
  axis: "rows" | "cols",
  at: number,
  count: number,
): AutoFilter | null {
  const range = shiftRange(filter.range, axis, at, count)
  if (!range) return null
  if (axis === "rows") return { ...filter, range }
  const { sort, ...rest } = filter
  const sortCol = sort ? filter.range.c0 + sort.col : -1
  const sortKept = sort && (count > 0 || sortCol < at || sortCol >= at - count)
  const columns: Record<number, ColumnFilter> = {}
  for (const [key, column] of Object.entries(filter.columns)) {
    const col = filter.range.c0 + Number(key)
    let moved = col
    if (count > 0) moved = col >= at ? col + count : col
    else if (col >= at && col < at - count) continue
    else if (col >= at - count) moved = col + count
    const offset = moved - range.c0
    if (offset >= 0 && moved <= range.c1) columns[offset] = column
  }
  const next: AutoFilter = { ...rest, range, columns }
  if (sort && sortKept) {
    const moved = sortCol >= at ? sortCol + count : sortCol
    next.sort = { ...sort, col: moved - range.c0 }
  }
  return next
}

/** One entry of a filter dropdown's value list: the text a cell shows, and whether it's ticked. */
export type FilterListItem = { text: string; checked: boolean }
/** A date of the list (shown in the year › month › day tree). */
export type FilterDateItem = { year: number; month: number; day: number; checked: boolean }

/** What a column's dropdown offers. */
export type FilterChoices = {
  /** Which submenu it shows: Text, Number or Date Filters (by what most of the column holds). */
  kind: "text" | "number" | "date"
  /** Values other than dates, in sort order (numbers, text, logicals, errors). */
  values: FilterListItem[]
  dates: FilterDateItem[]
  /** (Blanks), when the column has empty cells. */
  blanks: { checked: boolean } | null
  /** The fills and font colours in the column (null: none / automatic), for Sort by Color and Filter by Color. */
  cellColors: (string | null)[]
  fontColors: (string | null)[]
}

const typeRank = (v: unknown) =>
  typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3

/**
 * A column's dropdown: the values of the rows the other columns' filters leave (as Excel lists them), each ticked
 * when this column's own filter shows it.
 */
export function filterChoices(
  sheet: Sheet,
  styles: StyleTable,
  filter: AutoFilter,
  col: number,
  today: number,
): FilterChoices {
  const offset = col - filter.range.c0
  const others: AutoFilter = { range: filter.range, columns: { ...filter.columns } }
  delete others.columns[offset]
  const own = filter.columns[offset]
  const left = filterVisibility(sheet, styles, others, today)
  const passing = own
    ? filterVisibility(sheet, styles, { range: filter.range, columns: { [offset]: own } }, today)
    : null
  const cells = filterColumnCells(sheet, styles, filter.range, col)
  const values = new Map<string, { text: string; value: unknown; checked: boolean }>()
  const dates = new Map<number, FilterDateItem>()
  let blanks: { checked: boolean } | null = null
  let numbers = 0
  let dated = 0
  let texts = 0
  const fills = new Set<string | null>()
  const fonts = new Set<string | null>()
  cells.forEach((cell, i) => {
    if (!left[i]) return
    const checked = passing ? passing[i] : true
    const style = styles.get(sheet.styleAt(filter.range.r0 + 1 + i, col))
    fills.add(style.fill?.toUpperCase() ?? null)
    fonts.add(style.fontColor?.toUpperCase() ?? null)
    if (isBlank(cell.data)) {
      blanks = { checked: (blanks?.checked ?? false) || checked }
      return
    }
    const value = cell.data!.value
    if (typeof value === "number" && cell.code && isDateFormat(cell.code)) {
      dated++
      const day = Math.floor(value)
      const d = serialToDate(day)
      const item = dates.get(day) ?? { year: d.year, month: d.month, day: d.day, checked: false }
      item.checked ||= checked
      dates.set(day, item)
      return
    }
    if (typeof value === "number") numbers++
    else texts++
    const key = norm(cell.text)
    const item = values.get(key) ?? { text: cell.text, value, checked: false }
    item.checked ||= checked
    values.set(key, item)
  })
  const list = [...values.values()].sort((a, b) => {
    const rank = typeRank(a.value) - typeRank(b.value)
    if (rank) return rank
    if (typeof a.value === "number" && typeof b.value === "number") return a.value - b.value
    return compareSortText(norm(a.text), norm(b.text))
  })
  const kind = dated > numbers && dated >= texts ? "date" : numbers > texts ? "number" : "text"
  return {
    kind,
    values: list.map((v) => ({ text: v.text, checked: v.checked })),
    dates: [...dates.entries()].sort((a, b) => a[0] - b[0]).map(([, d]) => d),
    blanks,
    cellColors: [...fills],
    fontColors: [...fonts],
  }
}

/**
 * The values filter for what's ticked in a dropdown, or null when everything is (no filter). Ticked dates are kept
 * at the widest whole level, as Excel saves them: a whole year, a whole month, else single days.
 */
export function valuesFilter(choices: FilterChoices): ColumnFilter | null {
  const all =
    choices.values.every((v) => v.checked) && choices.dates.every((d) => d.checked) && (choices.blanks?.checked ?? true)
  if (all) return null
  const groups: DateGroup[] = []
  const years = new Map<number, FilterDateItem[]>()
  for (const d of choices.dates) years.set(d.year, [...(years.get(d.year) ?? []), d])
  for (const [year, days] of years) {
    if (days.every((d) => d.checked)) {
      groups.push({ year })
      continue
    }
    const months = new Map<number, FilterDateItem[]>()
    for (const d of days) months.set(d.month, [...(months.get(d.month) ?? []), d])
    for (const [month, inMonth] of months) {
      if (inMonth.every((d) => d.checked)) groups.push({ year, month })
      else for (const d of inMonth) if (d.checked) groups.push({ year, month, day: d.day })
    }
  }
  const out: ColumnFilter = { kind: "values", values: choices.values.filter((v) => v.checked).map((v) => v.text) }
  if (choices.blanks?.checked) out.blanks = true
  if (groups.length) out.dates = groups
  return out
}
