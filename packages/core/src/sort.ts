import type { Range } from "./address"
import type { CellData } from "./cell-store"
import { copyFormula } from "./formula/adjust"
import type { AnnotationWrite } from "./annotations"
import type { CellWrite, Op } from "./ops"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"
import type { CellValue } from "./values"

/**
 * Sorting as Excel does (checked against Excel 365 on Windows, en-US):
 * - ascending: numbers, then text, then FALSE and TRUE, then errors; blanks always last, in both orders;
 * - text compares as Windows' word sort: case doesn't matter (unless asked), accents do (a < ä), ligatures and ß
 *   expand (æ = ae, ß = ss), hyphens and apostrophes count only to break a tie (coop < co-op), symbols come before
 *   digits and digits before letters;
 * - equal keys keep their order (the sort is stable), descending too.
 */

// Windows' order of the symbols before the digits (from Excel sorting each one followed by "x").
const SYMBOLS = ` !"#$%&()*,./:;?@[\\]^_\`{|}~€+<=>§°`
const SYMBOL_WEIGHT = new Map([...SYMBOLS].map((c, i) => [c, i + 1] as const))
const OTHER_SYMBOL = 200
const DIGIT_BASE = 1000
const LETTER_BASE = 2000
const OTHER_SCRIPT = 0x20000

const EXPANSIONS: Record<string, string> = {
  æ: "ae",
  œ: "oe",
  ß: "ss",
  ﬁ: "fi",
  ﬂ: "fl",
  ﬀ: "ff",
  ĳ: "ij",
  þ: "th",
}
// Letters that aren't a base letter plus a combining mark: their base, and an accent weight.
const STROKED: Record<string, string> = { ø: "o", đ: "d", ł: "l", ħ: "h", ð: "d", ı: "i", ŧ: "t" }
const FRACTIONS: Record<string, number> = { "¼": 0.25, "½": 0.5, "¾": 0.75 }
const SUPERSCRIPTS: Record<string, number> = { "⁰": 0, "¹": 1, "²": 2, "³": 3 }

/** A string's collation key: primary weights, accents, case, and how many ignorable marks it had. */
type TextKey = { primary: number[]; accents: number[]; cases: number[]; ignorable: number }

export function textKey(text: string): TextKey {
  const primary: number[] = []
  const accents: number[] = []
  const cases: number[] = []
  let ignorable = 0
  for (const raw of text) {
    if (raw === "-" || raw === "'" || raw === "­" || raw === "’") {
      ignorable++
      continue
    }
    const lower = raw.toLowerCase()
    const upper = lower !== raw
    const expanded = EXPANSIONS[lower]
    if (expanded) {
      for (const c of expanded) {
        primary.push(LETTER_BASE + c.charCodeAt(0))
        accents.push(0)
        cases.push(upper ? 1 : 0)
      }
      continue
    }
    const decomposed = lower.normalize("NFD")
    let base = decomposed[0]
    let accent = 0
    for (let i = 1; i < decomposed.length; i++) accent = accent * 31 + decomposed.charCodeAt(i)
    if (STROKED[base]) {
      accent = accent * 31 + 0x338
      base = STROKED[base]
    }
    let weight: number
    if (base >= "0" && base <= "9") weight = DIGIT_BASE + (base.charCodeAt(0) - 48) * 4
    else if (raw in FRACTIONS) weight = DIGIT_BASE + FRACTIONS[raw] * 4
    else if (raw in SUPERSCRIPTS) {
      weight = DIGIT_BASE + SUPERSCRIPTS[raw] * 4
      accent = 1
    } else if (base >= "a" && base <= "z") weight = LETTER_BASE + base.charCodeAt(0)
    else if (SYMBOL_WEIGHT.has(raw)) weight = SYMBOL_WEIGHT.get(raw)!
    else if (/[\p{L}\p{N}]/u.test(base)) weight = OTHER_SCRIPT + (base.codePointAt(0) ?? 0)
    else weight = OTHER_SYMBOL + (raw.codePointAt(0) ?? 0) / 0x110000
    primary.push(weight)
    accents.push(accent)
    cases.push(upper ? 1 : 0)
  }
  return { primary, accents, cases, ignorable }
}

function compareArrays(a: number[], b: number[]): number {
  const n = Math.min(a.length, b.length)
  for (let i = 0; i < n; i++) if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1
  return a.length - b.length
}

/** Two texts in Excel's order: letters, then accents, then (when asked) case, lowercase first; then - and '. */
export function compareTextKeys(a: TextKey, b: TextKey, matchCase = false): number {
  return (
    compareArrays(a.primary, b.primary) ||
    compareArrays(a.accents, b.accents) ||
    (matchCase ? compareArrays(a.cases, b.cases) : 0) ||
    a.ignorable - b.ignorable
  )
}

export function compareSortText(a: string, b: string, matchCase = false): number {
  return compareTextKeys(textKey(a), textKey(b), matchCase)
}

const typeRank = (v: CellValue) =>
  typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3

/** One level of a sort: a column (or row, sorting left to right) and how to order it. */
export type SortKey = {
  /** The column (or row) in sheet coordinates. */
  index: number
  descending?: boolean
  /** Sort On: cell values (default), the cell's fill colour or its font colour. */
  on?: "values" | "cellColor" | "fontColor"
  /** For colour keys: the colour that goes first ("On Top") or last ("On Bottom"); null is "No Cell Color". */
  color?: string | null
  /** For value keys: a custom list ("Jan, Feb, Mar…"); its entries order before everything else, in list order. */
  list?: string[]
}

export type SortOptions = {
  keys: SortKey[]
  /** The first row (or column) holds headers and stays put. */
  header?: boolean
  matchCase?: boolean
  /** Sort left to right: columns move instead of rows. */
  byColumns?: boolean
}

export class SortError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "SortError"
  }
}

type Line = { index: number; keys: unknown[] }

/** A cell's sort key under one level: a value (with its key cached), a colour match, or a list position. */
function keyFor(data: CellData | undefined, key: SortKey, styles: StyleTable): unknown {
  if (key.on === "cellColor" || key.on === "fontColor") {
    const style = styles.get(data?.style ?? 0)
    const color = (key.on === "cellColor" ? style.fill : style.fontColor) ?? null
    return (color?.toUpperCase() ?? null) === (key.color?.toUpperCase() ?? null) ? 0 : 1
  }
  const value = data?.value
  if (value === undefined || value === "") return undefined
  if (key.list && typeof value === "string") {
    const at = key.list.findIndex((item) => compareSortText(item, value) === 0)
    if (at >= 0) return { list: at }
  }
  return typeof value === "string" ? { text: textKey(value) } : { value }
}

function compareKeys(a: unknown, b: unknown, key: SortKey, matchCase: boolean): number {
  if (key.on === "cellColor" || key.on === "fontColor") {
    // "On Bottom" is descending: the colour last.
    return key.descending ? (b as number) - (a as number) : (a as number) - (b as number)
  }
  // Blanks last, whichever the order.
  if (a === undefined || b === undefined) return a === undefined ? (b === undefined ? 0 : 1) : -1
  const ka = a as { list?: number; text?: TextKey; value?: CellValue }
  const kb = b as { list?: number; text?: TextKey; value?: CellValue }
  let c: number
  if (ka.list !== undefined || kb.list !== undefined) {
    // List entries first, in list order; then everything else in the usual order.
    if (ka.list !== undefined && kb.list !== undefined) c = ka.list - kb.list
    else c = ka.list !== undefined ? -1 : 1
    return key.descending ? -c : c
  }
  if (ka.text && kb.text) c = compareTextKeys(ka.text, kb.text, matchCase)
  else {
    const va = ka.text ? "" : (ka.value as CellValue)
    const vb = kb.text ? "" : (kb.value as CellValue)
    const ra = ka.text ? 1 : typeRank(va)
    const rb = kb.text ? 1 : typeRank(vb)
    if (ra !== rb) c = ra - rb
    else if (ra === 0) c = (va as number) - (vb as number)
    else if (ra === 2) c = Number(va) - Number(vb)
    else c = 0
  }
  return key.descending ? -c : c
}

/**
 * The operations that sort `range` on `sheet`: each row (or column) moves whole, formulas' relative references
 * shift with it as when copied, notes and comments go along. Hidden rows (columns) stay where they are, as in
 * Excel. Throws a SortError Excel would show (merged cells of different sizes).
 */
export function sortOps(sheetIndex: number, sheet: Sheet, styles: StyleTable, range: Range, o: SortOptions): Op[] {
  const byColumns = Boolean(o.byColumns)
  if (sheet.merges.overlapping(range).length)
    throw new SortError("To do this, all the merged cells need to be the same size.")
  const first = (byColumns ? range.c0 : range.r0) + (o.header ? 1 : 0)
  const last = byColumns ? range.c1 : range.r1
  const axis = byColumns ? sheet.cols : sheet.rows
  const lines: Line[] = []
  for (let i = first; i <= last; i++) {
    if (axis.isHidden(i)) continue
    lines.push({
      index: i,
      keys: o.keys.map((k) => keyFor(byColumns ? sheet.cells.get(k.index, i) : sheet.cells.get(i, k.index), k, styles)),
    })
  }
  const matchCase = Boolean(o.matchCase)
  const sorted = [...lines].sort((a, b) => {
    for (let k = 0; k < o.keys.length; k++) {
      const c = compareKeys(a.keys[k], b.keys[k], o.keys[k], matchCase)
      if (c) return c
    }
    return a.index - b.index
  })
  if (sorted.every((line, i) => line.index === lines[i].index)) return []

  const cells: CellWrite[] = []
  const notes: AnnotationWrite[] = []
  const across = byColumns ? [range.r0, range.r1] : [range.c0, range.c1]
  sorted.forEach((from, i) => {
    const to = lines[i].index
    if (from.index === to) return
    const shift = to - from.index
    for (let j = across[0]; j <= across[1]; j++) {
      const [sr, sc] = byColumns ? [j, from.index] : [from.index, j]
      const [tr, tc] = byColumns ? [j, to] : [to, j]
      const data = sheet.cells.get(sr, sc)
      cells.push({
        row: tr,
        col: tc,
        data: data ? moved(data, byColumns ? 0 : shift, byColumns ? shift : 0) : undefined,
      })
      const note = sheet.annotations.get(sr, sc) ?? null
      if (note || sheet.annotations.get(tr, tc)) notes.push({ row: tr, col: tc, value: note })
    }
  })
  const ops: Op[] = [{ kind: "setCells", sheet: sheetIndex, cells }]
  if (notes.length) ops.push({ kind: "annotations", sheet: sheetIndex, writes: notes })
  return ops
}

function moved(data: CellData, dr: number, dc: number): CellData | undefined {
  if (data.spill) return data.style ? { style: data.style } : undefined
  const copy: CellData = { ...data }
  if (copy.input?.startsWith("=")) copy.input = copyFormula(copy.input, dr, dc)
  return copy
}

/**
 * Whether a range's first row looks like headers, as Excel guesses for Sort: every filled cell in it is text, and
 * under some of them is a value of another kind (a number under "Amount") or the row is formatted differently (bold).
 */
export function guessHeader(sheet: Sheet, styles: StyleTable, range: Range, byColumns = false): boolean {
  const lines = byColumns ? range.c1 - range.c0 : range.r1 - range.r0
  if (lines < 1) return false
  let other = false
  let styled = false
  let filled = 0
  const across = byColumns ? [range.r0, range.r1] : [range.c0, range.c1]
  for (let j = across[0]; j <= across[1]; j++) {
    const head = byColumns ? sheet.cells.get(j, range.c0) : sheet.cells.get(range.r0, j)
    const next = byColumns ? sheet.cells.get(j, range.c0 + 1) : sheet.cells.get(range.r0 + 1, j)
    if (head?.value === undefined || head.value === "") continue
    if (typeof head.value !== "string" || head.input?.startsWith("=")) return false
    filled++
    if (next?.value !== undefined && next.value !== "" && typeof next.value !== "string") other = true
    const a = styles.get(head.style ?? 0)
    const b = styles.get(next?.style ?? 0)
    if (Boolean(a.bold) !== Boolean(b.bold) || a.fill !== b.fill || a.fontColor !== b.fontColor) styled = true
  }
  return filled > 0 && (other || styled)
}
