import type { CellData } from "./cell-store"
import type { CellValue } from "./values"

/** Lists the fill handle continues (Excel's built-in custom lists, en-US). */
const LISTS: string[][] = [
  ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
  ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ],
]

type Item =
  | { kind: "number"; value: number }
  | { kind: "list"; list: number; index: number; text: string }
  | { kind: "numbered"; prefix: string; number: number; digits: number; suffix: string; ordinal: boolean }
  | { kind: "other" }

// A number at the start of text ("1st", "2 x 3") counts first; otherwise one at the end ("Item 7", "Q3"). A number in
// the middle ("a1b") is not counted. All as Excel's AutoFill (checked against Excel).
const LEADING_RE = /^(\d+)(\D.*)$/
const TRAILING_RE = /^(.*\D)(\d+)$/
const ORDINAL_RE = /^(st|nd|rd|th)$/i
const QUARTER_RE = /^(q|qtr|quarter)\s?$/i

function classify(data: CellData | undefined): Item {
  const value = data?.value
  if (typeof value === "number") {
    return { kind: "number", value }
  }
  if (typeof value !== "string" || data?.input?.startsWith("=")) {
    return { kind: "other" }
  }
  for (let list = 0; list < LISTS.length; list++) {
    const index = LISTS[list].findIndex((name) => name.toLowerCase() === value.toLowerCase())
    if (index !== -1) {
      return { kind: "list", list, index, text: value }
    }
  }
  const leading = LEADING_RE.exec(value)
  if (leading && leading[1].length <= 15) {
    const suffix = leading[2]
    return {
      kind: "numbered",
      prefix: "",
      number: Number(leading[1]),
      digits: leading[1].length,
      suffix,
      ordinal: ORDINAL_RE.test(suffix),
    }
  }
  const trailing = TRAILING_RE.exec(value)
  if (trailing && trailing[2].length <= 15) {
    return {
      kind: "numbered",
      prefix: trailing[1],
      number: Number(trailing[2]),
      digits: trailing[2].length,
      suffix: "",
      ordinal: false,
    }
  }
  return { kind: "other" }
}

/** Keeps the case style of the typed name: "MON" → "TUE", "mon" → "tue", "Mon" → "Tue". */
function matchCase(template: string, name: string) {
  if (template === template.toUpperCase()) return name.toUpperCase()
  if (template === template.toLowerCase()) return name.toLowerCase()
  return name
}

function ordinalSuffix(n: number, template: string) {
  const mod100 = n % 100
  const suffix = mod100 >= 11 && mod100 <= 13 ? "th" : (["th", "st", "nd", "rd"][n % 10] ?? "th")
  return template === template.toUpperCase() ? suffix.toUpperCase() : suffix
}

/** Least-squares line through (0, y0), (1, y1), …: Excel's AutoFill trend for two or more numbers. */
function trend(values: number[]): { start: number; step: number } {
  const n = values.length
  const meanX = (n - 1) / 2
  const meanY = values.reduce((sum, y) => sum + y, 0) / n
  let sxy = 0
  let sxx = 0
  values.forEach((y, x) => {
    sxy += (x - meanX) * (y - meanY)
    sxx += (x - meanX) ** 2
  })
  const step = sxx === 0 ? 0 : sxy / sxx
  return { start: meanY - step * meanX, step }
}

function clean(value: number) {
  return Number(value.toPrecision(15))
}

/** Value at position x of a run (x = 0 is the run's first cell), or null to copy the run. */
type RunSeries = ((x: number) => CellValue) | null

/** Consecutive source cells of the same kind; Excel continues each run as its own series. */
type Run = { start: number; items: Item[] }

function sameRun(a: Item, b: Item) {
  if (a.kind !== b.kind) return false
  if (a.kind === "list" && b.kind === "list") return a.list === b.list
  if (a.kind === "numbered" && b.kind === "numbered") return a.prefix === b.prefix && a.suffix === b.suffix
  return true
}

function splitRuns(items: Item[]): Run[] {
  const runs: Run[] = []
  items.forEach((item, index) => {
    const last = runs.at(-1)
    if (last && sameRun(last.items[0], item)) {
      last.items.push(item)
    } else {
      runs.push({ start: index, items: [item] })
    }
  })
  return runs
}

/**
 * The series of one run. `alone` is true when the run is the whole selection: a lone number then copies (Ctrl
 * counts), while inside a mixed selection it counts by 1, as Excel does.
 */
function runSeries(items: Item[], alone: boolean, invert: boolean): RunSeries {
  const first = items[0]
  const m = items.length
  if (first.kind === "number") {
    const asSeries = alone && m === 1 ? invert : !invert
    if (!asSeries) return null
    const values = items.map((item) => (item as { value: number }).value)
    const { start, step } = m === 1 ? { start: values[0], step: 1 } : trend(values)
    return (x) => clean(start + step * x)
  }
  if (invert || first.kind === "other") {
    return null
  }
  if (first.kind === "list") {
    const list = LISTS[first.list]
    const indexes = items.map((item) => (item as { index: number }).index)
    const step = m === 1 ? 1 : (indexes[1] - indexes[0] + list.length) % list.length
    const even = indexes.every((index, i) => i === 0 || (indexes[i - 1] + step) % list.length === index)
    if (!even) return null
    return (x) => matchCase(first.text, list[(((first.index + step * x) % list.length) + list.length) % list.length])
  }
  const numbers = items.map((item) => (item as { number: number }).number)
  const step = m === 1 ? 1 : numbers[1] - numbers[0]
  // Text numbers only continue an even step; anything else copies.
  if (!numbers.every((n, i) => i === 0 || n - numbers[i - 1] === step)) return null
  const quarter = QUARTER_RE.test(first.prefix) && !first.suffix && numbers.every((n) => n >= 1 && n <= 4)
  return (x) => {
    let number = numbers[0] + step * x
    if (quarter) {
      number = ((((number - 1) % 4) + 4) % 4) + 1
    }
    // Text series never go negative: Excel counts back up past zero ("Item 0", "Item 1", …).
    number = Math.abs(number)
    const text = String(number).padStart(first.digits, "0")
    const suffix = first.ordinal ? ordinalSuffix(number, first.suffix) : first.suffix
    return `${first.prefix}${text}${suffix}`
  }
}

/**
 * What the fill handle writes when dragged `count` cells beyond `source` (the selected cells along the drag, in
 * sheet order). `dir` is 1 when dragging down/right, -1 up/left. `invert` is Ctrl held while dragging: it turns a
 * copy into a series and a series into a copy. Results are in order of distance from the source. Formats repeat
 * with the source's pattern.
 *
 * The source splits into runs of consecutive cells of the same kind (numbers; text with the same words around a
 * number; names from one list; anything else), and each run continues as its own series every time the pattern
 * repeats: "x, 1, 2" fills "x, 3, 4, x, 5, 6".
 */
export function fillSeries(
  source: readonly (CellData | undefined)[],
  count: number,
  dir: 1 | -1,
  invert = false,
): (CellData | undefined)[] {
  const n = source.length
  const items = source.map(classify)
  const runs = splitRuns(items)
  const series = runs.map((run) => runSeries(run.items, runs.length === 1, invert))
  const runOf = new Array<number>(n)
  runs.forEach((run, r) => run.items.forEach((_, i) => (runOf[run.start + i] = r)))

  const out: (CellData | undefined)[] = []
  for (let k = 1; k <= count; k++) {
    // Position on the source's axis: n-1+k going forward, -k going back.
    const p = dir === 1 ? n - 1 + k : -k
    const j = ((p % n) + n) % n
    const cycle = Math.floor(p / n)
    const r = runOf[j]
    const run = runs[r]
    const style = source[j]?.style
    let data: CellData | undefined
    const make = series[r]
    if (make) {
      data = { value: make(cycle * run.items.length + (j - run.start)) }
    } else {
      const copied = source[j]
      data = copied && (copied.value !== undefined || copied.input !== undefined) ? { ...copied } : undefined
    }
    if (style) {
      data = { ...(data ?? {}), style }
    } else if (data) {
      delete data.style
    }
    out.push(data && Object.keys(data).length ? data : undefined)
  }
  return out
}
