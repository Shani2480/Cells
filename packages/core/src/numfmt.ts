import { serialToDate, serialWeekday } from "./formula/dates"
import { roundDecimal } from "./rounding"

/**
 * Excel's number-format language: up to four sections (positive; negative; zero; text), conditions ([>=100]),
 * colours ([Red]), literals ("x", \x), spacing (_x) and fill (*x), digits (0 # ?), thousands separators and scaling
 * (,), percent, scientific (E+), fractions (# ?/?), dates and times (y m d h s AM/PM, elapsed [h]) and text (@).
 * en-US separators. Used by the TEXT function now and by cell formatting in Phase 3.
 */

/** `overflow`: a date or time out of Excel's range (the cell shows ####, TEXT gives #VALUE!). */
export type FormatResult = { text: string; color?: string; overflow?: true }

/**
 * In cell mode (`formatWithCode(…, { cell: true })`), `_x` comes out as one private-use character standing for "a gap
 * as wide as x" and `*x` as one standing for "x repeated to fill the cell". One character each, so the digit
 * formatting (which reverses the integer part) can't split them. The grid's painter reads them with `formatMark`;
 * TEXT has no width, so it shows a space and nothing.
 */
const PAD_BASE = 0xe100
const FILL_BASE = 0xe200

const markOf = (base: number, ch: string | undefined) => {
  const code = ch?.charCodeAt(0) ?? 32
  return String.fromCharCode(base + (code < 128 ? code : 32))
}

/** What a cell-mode character means: a gap as wide as `ch`, the fill character `ch`, or plain text. */
export function formatMark(c: string): { kind: "pad" | "fill"; ch: string } | null {
  const code = c.charCodeAt(0)
  if (code >= PAD_BASE && code < PAD_BASE + 128) return { kind: "pad", ch: String.fromCharCode(code - PAD_BASE) }
  if (code >= FILL_BASE && code < FILL_BASE + 128) return { kind: "fill", ch: String.fromCharCode(code - FILL_BASE) }
  return null
}

/** Whether cell-mode text has gaps or a fill to lay out. */
export const hasFormatMarks = (text: string) => /[\ue100-\ue27f]/.test(text)

type Section = { tokens: Tok[]; condition?: (n: number) => boolean; color?: string; hasCondition: boolean }

type Tok =
  | { t: "lit"; s: string }
  | { t: "digit"; c: "0" | "#" | "?" }
  | { t: "point" }
  | { t: "comma" }
  | { t: "percent" }
  | { t: "exp"; sign: "+" | "-" }
  | { t: "slash" }
  | { t: "text" } // @
  | { t: "general" }
  | { t: "date"; code: string } // y, yy, yyyy, m, mm, mmm, mmmm, mmmmm, d, dd, ddd, dddd, h, hh, s, ss, AM/PM, A/P, [h], [m], [s], .0 (fraction of second)

const cache = new Map<string, Section[]>()

function parse(code: string, cell = false): Section[] {
  const key = cell ? "\u0001" + code : code
  let sections = cache.get(key)
  if (sections) return sections
  sections = splitSections(code).map((src) => parseSection(src, cell))
  if (cache.size > 500) cache.clear()
  cache.set(key, sections)
  return sections
}

function splitSections(code: string): string[] {
  const out: string[] = []
  let cur = ""
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      const j = end < 0 ? code.length : end
      cur += code.slice(i, j + 1)
      i = j
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      cur += code.slice(i, i + 2)
      i++
    } else if (ch === "[") {
      const end = code.indexOf("]", i)
      const j = end < 0 ? code.length : end
      cur += code.slice(i, j + 1)
      i = j
    } else if (ch === ";") {
      out.push(cur)
      cur = ""
    } else cur += ch
  }
  out.push(cur)
  return out
}

const COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#00FF00",
  blue: "#0000FF",
  yellow: "#FFFF00",
  magenta: "#FF00FF",
  cyan: "#00FFFF",
}

function parseSection(src: string, cell = false): Section {
  const tokens: Tok[] = []
  const section: Section = { tokens, hasCondition: false }
  let i = 0
  const lower = src.toLowerCase()
  while (i < src.length) {
    const ch = src[i]
    const lc = lower[i]
    if (ch === '"') {
      const end = src.indexOf('"', i + 1)
      const j = end < 0 ? src.length : end
      tokens.push({ t: "lit", s: src.slice(i + 1, j) })
      i = j + 1
      continue
    }
    if (ch === "\\") {
      tokens.push({ t: "lit", s: src[i + 1] ?? "" })
      i += 2
      continue
    }
    if (ch === "_") {
      tokens.push({ t: "lit", s: cell ? markOf(PAD_BASE, src[i + 1]) : " " })
      i += 2
      continue
    }
    if (ch === "*") {
      // Fill to the column width: in TEXT (no width) it shows nothing.
      if (cell) tokens.push({ t: "lit", s: markOf(FILL_BASE, src[i + 1]) })
      i += 2
      continue
    }
    if (ch === "[") {
      const end = src.indexOf("]", i)
      const inner = src.slice(i + 1, end < 0 ? src.length : end)
      i = end < 0 ? src.length : end + 1
      const il = inner.toLowerCase()
      if (/^(h+|m+|s+)$/.test(il)) tokens.push({ t: "date", code: `[${il[0]}]${il.length > 1 ? "2" : ""}` })
      else if (COLORS[il]) section.color = COLORS[il]
      else if (/^color\s*\d+$/.test(il)) section.color = il
      else if (/^(<=|>=|<>|<|>|=)\s*-?[\d.]+(e[+-]?\d+)?$/.test(il)) {
        const m = /^(<=|>=|<>|<|>|=)\s*(.+)$/.exec(il)!
        const v = Number(m[2])
        const op = m[1]
        section.hasCondition = true
        section.condition = (n) =>
          op === "<"
            ? n < v
            : op === ">"
              ? n > v
              : op === "<="
                ? n <= v
                : op === ">="
                  ? n >= v
                  : op === "="
                    ? n === v
                    : n !== v
      } else if (inner.startsWith("$")) {
        // [$€-407]: a currency symbol and a locale; show the symbol.
        tokens.push({ t: "lit", s: inner.slice(1).split("-")[0] })
      }
      continue
    }
    if (lower.startsWith("general", i)) {
      tokens.push({ t: "general" })
      i += 7
      continue
    }
    if (lower.startsWith("am/pm", i)) {
      tokens.push({ t: "date", code: "AM/PM" })
      i += 5
      continue
    }
    if (lower.startsWith("a/p", i)) {
      tokens.push({ t: "date", code: src.slice(i, i + 3) })
      i += 3
      continue
    }
    if ("ymdhse".includes(lc) && !(lc === "e" && /e[+-]/.test(lower.slice(i, i + 2)))) {
      let j = i
      while (j < src.length && lower[j] === lc) j++
      tokens.push({ t: "date", code: lower.slice(i, j) })
      i = j
      continue
    }
    if (lc === "e" && (src[i + 1] === "+" || src[i + 1] === "-")) {
      tokens.push({ t: "exp", sign: src[i + 1] as "+" | "-" })
      i += 2
      continue
    }
    if (ch === "0" || ch === "#" || ch === "?") tokens.push({ t: "digit", c: ch })
    else if (ch === ".") tokens.push({ t: "point" })
    else if (ch === ",") tokens.push({ t: "comma" })
    else if (ch === "%") tokens.push({ t: "percent" })
    else if (ch === "/") tokens.push({ t: "slash" })
    else if (ch === "@") tokens.push({ t: "text" })
    else tokens.push({ t: "lit", s: ch })
    i++
  }
  // ".0" after seconds is a fraction of a second.
  for (let k = 0; k < tokens.length; k++) {
    const tk = tokens[k]
    if (tk.t === "point" && tokens[k - 1]?.t === "date" && /s/.test((tokens[k - 1] as { code: string }).code)) {
      let n = 0
      while (tokens[k + 1 + n]?.t === "digit" && (tokens[k + 1 + n] as { c: string }).c === "0") n++
      if (n) tokens.splice(k, n + 1, { t: "date", code: `.${"0".repeat(n)}` })
    }
  }
  return section
}

function isDateSection(s: Section) {
  return s.tokens.some((t) => t.t === "date")
}

/** Formats a number or text with a format code. */
export function formatWithCode(
  value: number | string | boolean,
  code: string,
  options: { cell?: boolean } = {},
): FormatResult {
  const sections = parse(code, options.cell)
  if (typeof value === "boolean") return { text: value ? "TRUE" : "FALSE" }
  if (typeof value === "string") {
    const textSection =
      sections[3] ?? (sections.length === 1 && sections[0].tokens.some((t) => t.t === "text") ? sections[0] : null)
    if (!textSection) return { text: value }
    return {
      text: textSection.tokens.map((t) => (t.t === "text" ? value : t.t === "lit" ? t.s : "")).join(""),
      color: textSection.color,
    }
  }
  const { section, negativeHasOwn } = pickSection(sections, value)
  const v = negativeHasOwn ? Math.abs(value) : value
  if (section.tokens.length === 0) return { text: "", color: section.color }
  if (section.tokens.some((t) => t.t === "text") && !section.tokens.some((t) => t.t === "digit")) {
    // @ with a number: the number as General text ("@" alone shows 1234.5678 as it is).
    const n = String(Number(roundDecimal(v, 15 - Math.max(1, Math.floor(Math.log10(Math.abs(v) || 1)) + 1))))
    return {
      text: section.tokens.map((t) => (t.t === "text" ? n : t.t === "lit" ? t.s : "")).join(""),
      color: section.color,
    }
  }
  if (isDateSection(section)) {
    if (value < 0 || value > 2_958_465.99999) return { text: "#".repeat(10), overflow: true }
    return { text: formatDate(v, section.tokens), color: section.color }
  }
  return { text: formatNumberTokens(v, section.tokens), color: section.color }
}

function pickSection(sections: Section[], n: number): { section: Section; negativeHasOwn: boolean } {
  const numeric = sections.slice(0, 3)
  if (numeric.some((s) => s.hasCondition)) {
    // With conditions: the first two sections test them; the third is "everything else".
    for (let i = 0; i < Math.min(2, numeric.length); i++) {
      const s = numeric[i]
      // A section picked by its condition shows no minus sign ([<0]"neg" shows "neg").
      if (s.condition?.(n)) return { section: s, negativeHasOwn: n < 0 }
    }
    // The "everything else" section: with two sections it keeps the sign ([>=1000]0,"K";0 shows -1235).
    const rest = numeric[2] ?? numeric[1]
    if (rest && !rest.hasCondition) return { section: rest, negativeHasOwn: numeric.length > 2 && n < 0 }
    return { section: { tokens: [{ t: "lit", s: "#".repeat(10) }], hasCondition: false }, negativeHasOwn: false }
  }
  if (n > 0 || numeric.length === 1) return { section: numeric[0], negativeHasOwn: false }
  if (n < 0) return { section: numeric[1], negativeHasOwn: true }
  return { section: numeric[2] ?? numeric[0], negativeHasOwn: false }
}

// ── Numbers ────────────────────────────────────────────────────────────────────────────────────────────────────

function formatNumberTokens(value: number, tokens: Tok[]): string {
  if (tokens.some((t) => t.t === "general")) {
    return tokens
      .map((t) => (t.t === "general" ? generalText(value) : t.t === "lit" ? t.s : t.t === "percent" ? "%" : ""))
      .join("")
  }
  const hasDigits = tokens.some((t) => t.t === "digit")
  if (!hasDigits) {
    // A format with no digit placeholders shows only its literals ("-" for zeros in accounting formats).
    return (value < 0 ? "-" : "") + tokens.map((t) => (t.t === "lit" ? t.s : t.t === "percent" ? "%" : "")).join("")
  }
  let v = value
  const percentCount = tokens.filter((t) => t.t === "percent").length
  v *= 100 ** percentCount

  // Thousands scaling: commas right before the point, or right after the last digit placeholder ("0.0,," scales
  // by a million).
  const lastDigit = lastIndex(tokens, (t) => t.t === "digit")
  const pointIndex = tokens.findIndex((t) => t.t === "point")
  const intEnd = pointIndex >= 0 ? pointIndex : lastDigit + 1
  const scaling = new Set<number>()
  for (let k = intEnd - 1; k >= 0 && tokens[k].t === "comma"; k--) scaling.add(k)
  for (let k = lastDigit + 1; k < tokens.length && tokens[k].t === "comma"; k++) scaling.add(k)
  const scale = scaling.size
  v /= 1000 ** scale
  const grouping = tokens.some(
    (t, k) =>
      t.t === "comma" &&
      k < intEnd &&
      tokens.slice(0, k).some((x) => x.t === "digit") &&
      tokens.slice(k + 1, intEnd).some((x) => x.t === "digit"),
  )

  const slash = tokens.findIndex((t) => t.t === "slash")
  if (slash >= 0) return formatFraction(v, tokens, slash, grouping)
  const expIndex = tokens.findIndex((t) => t.t === "exp")
  if (expIndex >= 0) return formatScientific(v, tokens, expIndex)

  const intTokens = tokens.slice(0, pointIndex >= 0 ? pointIndex : tokens.length)
  const fracTokens = pointIndex >= 0 ? tokens.slice(pointIndex + 1) : []
  const fracDigits = fracTokens.filter((t) => t.t === "digit").length
  const negative = v < 0
  const rounded = roundHalfAway(Math.abs(v), fracDigits)
  let [intPart, fracPart = ""] = rounded.split(".")
  fracPart = fracPart.padEnd(fracDigits, "0")
  if (intPart === "0") intPart = ""
  const showMinus = negative && (Number(intPart || "0") !== 0 || /[1-9]/.test(fracPart))

  const intText = fillInteger(intPart, intTokens, grouping)
  const fracText = fillFraction(fracPart, fracTokens)
  return (showMinus ? "-" : "") + intText + (pointIndex >= 0 ? "." + fracText : "")
}

function lastIndex<T>(arr: T[], f: (x: T) => boolean) {
  for (let i = arr.length - 1; i >= 0; i--) if (f(arr[i])) return i
  return -1
}

/** Rounds half away from zero on the 15-digit decimal value (as Excel displays) and writes it without exponent. */
function roundHalfAway(x: number, digits: number): string {
  const rounded = roundDecimal(x, digits)
  return rounded.toLocaleString("en-US", {
    useGrouping: false,
    minimumFractionDigits: digits,
    maximumFractionDigits: Math.min(digits, 100),
  })
}

/**
 * Puts integer digits into the placeholders from the right: 0 shows a zero where there is no digit, ? a space, #
 * nothing; digits beyond the placeholders all go into the leftmost one. Literals stay where they are.
 */
function fillInteger(digits: string, tokens: Tok[], grouping: boolean): string {
  const placeholders = tokens.filter((t) => t.t === "digit").length
  let di = digits.length - 1
  let seen = 0
  const result: string[] = []
  for (let k = tokens.length - 1; k >= 0; k--) {
    const t = tokens[k]
    if (t.t === "digit") {
      seen++
      result.push(di >= 0 ? digits[di--] : t.c === "0" ? "0" : t.c === "?" ? " " : "")
      if (seen === placeholders) while (di >= 0) result.push(digits[di--])
    } else if (t.t === "lit") result.push(t.s)
    else if (t.t === "percent") result.push("%")
  }
  // The pieces were collected right to left; each piece reads left to right already.
  const text = result.reverse().join("")
  return grouping ? groupDigits(text) : text
}

/** Inserts thousands separators into the digit run of the integer text. */
function groupDigits(text: string): string {
  return text.replace(/\d+/, (run) => run.replace(/\B(?=(\d{3})+(?!\d))/g, ","))
}

function fillFraction(digits: string, tokens: Tok[]): string {
  let di = 0
  const out: string[] = []
  // Trailing zeros: # drops them, ? turns them into spaces, 0 keeps them.
  const placeholders = tokens.filter((t) => t.t === "digit") as { c: string }[]
  const shown = [...digits]
  for (let k = placeholders.length - 1; k >= 0; k--) {
    if (placeholders[k].c === "0" || shown[k] !== "0") break
    shown[k] = placeholders[k].c === "?" ? " " : ""
  }
  for (const t of tokens) {
    if (t.t === "digit") {
      out.push(shown[di] ?? "")
      di++
    } else if (t.t === "lit") out.push(t.s)
    else if (t.t === "percent") out.push("%")
  }
  return out.join("")
}

function formatScientific(v: number, tokens: Tok[], expIndex: number): string {
  const mantissaTokens = tokens.slice(0, expIndex)
  const expTokens = tokens.slice(expIndex + 1)
  const point = mantissaTokens.findIndex((t) => t.t === "point")
  const intPlaceholders = (point >= 0 ? mantissaTokens.slice(0, point) : mantissaTokens).filter(
    (t) => t.t === "digit",
  ).length
  const fracPlaceholders = point >= 0 ? mantissaTokens.slice(point + 1).filter((t) => t.t === "digit").length : 0
  const negative = v < 0
  const abs = Math.abs(v)
  let exp = abs === 0 ? 0 : Math.floor(Math.log10(abs))
  // With more than one integer placeholder the exponent is a multiple of their count (engineering style).
  const step = Math.max(1, intPlaceholders)
  if (step > 1) exp = Math.floor(exp / step) * step
  else exp = exp - (intPlaceholders - 1)
  let mant = abs / 10 ** exp
  let text = mant.toFixed(fracPlaceholders)
  if (Number(text) >= 10 ** step) {
    exp += step
    mant = abs / 10 ** exp
    text = mant.toFixed(fracPlaceholders)
  }
  const [whole, f = ""] = text.split(".")
  // Excel's quirk: zero in an engineering format fills every integer placeholder (##0.0E+0 shows 000.0E+0).
  const i = abs === 0 && step > 1 ? "0".repeat(intPlaceholders) : whole
  const intTokens = point >= 0 ? mantissaTokens.slice(0, point) : mantissaTokens
  const mantText =
    fillInteger(i, intTokens, false) + (point >= 0 ? "." + fillFraction(f, mantissaTokens.slice(point + 1)) : "")
  const expDigits = expTokens.filter((t) => t.t === "digit").length || 1
  const expSign = exp < 0 ? "-" : (tokens[expIndex] as { sign: string }).sign === "+" ? "+" : ""
  return (negative ? "-" : "") + mantText + "E" + expSign + String(Math.abs(exp)).padStart(expDigits, "0")
}

function formatFraction(v: number, tokens: Tok[], slash: number, grouping: boolean): string {
  const negative = v < 0
  const abs = Math.abs(v)
  let denomTokens = tokens.slice(slash + 1)
  // A fixed denominator ("?/8", "?/100"): its digits parse as literals and 0 placeholders.
  let fixedText = ""
  let k = 0
  for (; k < denomTokens.length; k++) {
    const t = denomTokens[k]
    if (t.t === "lit" && /^\d$/.test(t.s)) fixedText += t.s
    else if (t.t === "digit" && t.c === "0" && fixedText) fixedText += "0"
    else break
  }
  const fixedDenom = /^[1-9]\d*$/.test(fixedText) ? Number(fixedText) : null
  if (fixedDenom) denomTokens = denomTokens.slice(k)
  const denomDigits = denomTokens.filter((t) => t.t === "digit") as { c: string }[]
  // Integer part: digit placeholders before the numerator, separated by a space literal.
  const before = tokens.slice(0, slash)
  let splitAt = -1
  for (let j = before.length - 1; j >= 0; j--) {
    if (before[j].t === "lit" && (before[j] as { s: string }).s.trim() === "") {
      splitAt = j
      break
    }
  }
  const hasWhole = splitAt > 0 && before.slice(0, splitAt).some((t) => t.t === "digit")
  const numTokens = hasWhole ? before.slice(splitAt + 1) : before
  const whole = hasWhole ? Math.floor(abs) : 0
  const frac = abs - whole
  let bestN = 0
  let bestD = 1
  if (fixedDenom) {
    bestD = fixedDenom
    bestN = Math.round(frac * fixedDenom)
  } else {
    const maxDen = 10 ** denomDigits.length - 1
    let bestErr = Infinity
    for (let d = 1; d <= maxDen; d++) {
      const n = Math.round(frac * d)
      const err = Math.abs(frac - n / d)
      if (err < bestErr - 1e-12) {
        bestErr = err
        bestN = n
        bestD = d
      }
    }
  }
  let w = whole
  if (bestN === bestD && hasWhole) {
    w += 1
    bestN = 0
  }
  const sign = negative ? "-" : ""
  const wholeText = hasWhole ? fillInteger(w === 0 ? "" : String(w), before.slice(0, splitAt), grouping) : ""
  const numText = fillInteger(String(bestN), numTokens, false)
  const denText = fixedDenom
    ? String(fixedDenom)
    : String(bestD) + " ".repeat(Math.max(0, denomDigits.filter((d) => d.c === "?").length - String(bestD).length))
  if (hasWhole && bestN === 0) {
    // No fraction: the whole number, then spaces as wide as the fraction would be when it has ? placeholders.
    const padded = [...numTokens, ...denomTokens].some((t) => t.t === "digit" && t.c === "?")
    const width =
      1 + numTokens.filter((t) => t.t === "digit").length + 1 + (fixedDenom ? fixedText.length : denomDigits.length)
    return sign + (wholeText || "0") + (padded ? " ".repeat(width) : "")
  }
  // With no whole number the separator stays as a space only when the numerator pads with ?.
  const firstNum = numTokens.find((t) => t.t === "digit") as { c: string } | undefined
  const separator = !hasWhole ? "" : wholeText || firstNum?.c === "?" ? " " : ""
  return sign + wholeText + separator + numText + "/" + denText
}

// ── Dates ──────────────────────────────────────────────────────────────────────────────────────────────────────

const MONTH_NAMES = [
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
]
const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

function formatDate(serial: number, tokens: Tok[]): string {
  const fracCode = tokens.find((t) => t.t === "date" && t.code.startsWith(".")) as { code: string } | undefined
  const fracDigits = fracCode ? fracCode.code.length - 1 : 0
  // Round to the shown precision: whole seconds unless the format shows fractions.
  // Whole units (seconds, or tenths… with .0) so 13:37:37.92 rounds to :38 without float drift.
  const perSecond = 10 ** fracDigits
  const unit = 86_400 * perSecond
  const units = Math.round(serial * unit)
  const day = Math.floor(units / unit)
  const dayUnits = units - day * unit
  const date = serialToDate(day)
  const secondsOfDay = Math.floor(dayUnits / perSecond)
  const hour = Math.floor(secondsOfDay / 3600)
  const minute = Math.floor((secondsOfDay % 3600) / 60)
  const second = secondsOfDay % 60
  const fracUnits = dayUnits % perSecond
  const elapsedSeconds = Math.floor(units / perSecond)
  const ampm = tokens.some((t) => t.t === "date" && (t.code === "AM/PM" || /^a\/p$/i.test(t.code)))
  const out: string[] = []
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k]
    if (t.t === "lit") {
      out.push(t.s)
      continue
    }
    if (t.t === "percent") {
      out.push("%")
      continue
    }
    if (t.t !== "date") {
      if (t.t === "digit") out.push(t.c === "0" ? "0" : "")
      else if (t.t === "point") out.push(".")
      else if (t.t === "comma") out.push(",")
      else if (t.t === "slash") out.push("/")
      continue
    }
    const c = t.code
    if (c.startsWith("[")) {
      const unitCode = c[1]
      const two = c.endsWith("2")
      const v =
        unitCode === "h"
          ? Math.floor(elapsedSeconds / 3600)
          : unitCode === "m"
            ? Math.floor(elapsedSeconds / 60)
            : elapsedSeconds
      out.push(String(v).padStart(two ? 2 : 1, "0"))
      continue
    }
    if (c.startsWith(".")) {
      out.push("." + String(fracUnits).padStart(fracDigits, "0"))
      continue
    }
    if (c === "AM/PM") {
      out.push(hour < 12 ? "AM" : "PM")
      continue
    }
    if (/^a\/p$/i.test(c)) {
      const pm = hour >= 12
      out.push(c[0] === "a" ? (pm ? "p" : "a") : pm ? "P" : "A")
      continue
    }
    const kind = c[0]
    const len = c.length
    if (kind === "y") out.push(len <= 2 ? String(date.year % 100).padStart(2, "0") : String(date.year))
    else if (kind === "d") {
      if (len === 1) out.push(String(date.day))
      else if (len === 2) out.push(String(date.day).padStart(2, "0"))
      else {
        const name = DAY_NAMES[serialWeekday(day)]
        out.push(len === 3 ? name.slice(0, 3) : name)
      }
    } else if (kind === "h") {
      const h = ampm ? hour % 12 || 12 : hour
      out.push(len >= 2 ? String(h).padStart(2, "0") : String(h))
    } else if (kind === "s") out.push(len >= 2 ? String(second).padStart(2, "0") : String(second))
    else if (kind === "m") {
      if (isMinute(tokens, k)) out.push(len >= 2 ? String(minute).padStart(2, "0") : String(minute))
      else if (len === 1) out.push(String(date.month))
      else if (len === 2) out.push(String(date.month).padStart(2, "0"))
      else if (len === 3) out.push(MONTH_NAMES[date.month - 1].slice(0, 3))
      else if (len === 5) out.push(MONTH_NAMES[date.month - 1][0])
      else out.push(MONTH_NAMES[date.month - 1])
    } else if (kind === "e") out.push(String(date.year))
  }
  return out.join("")
}

/** "m" is minutes right after an hour code or right before a seconds code (literals in between don't count). */
function isMinute(tokens: Tok[], k: number): boolean {
  if ((tokens[k] as { code: string }).code.length > 2) return false
  for (let j = k - 1; j >= 0; j--) {
    const t = tokens[j]
    if (t.t === "date") {
      if (/^h|^\[h/.test(t.code)) return true
      break
    }
  }
  for (let j = k + 1; j < tokens.length; j++) {
    const t = tokens[j]
    if (t.t === "date") return /^s|^\[s/.test(t.code)
  }
  return false
}

/** General inside a format code: up to 11 characters like a default-width cell. */
function generalText(n: number): string {
  if (n === 0) return "0"
  const abs = Math.abs(n)
  if (abs >= 1e11 || abs < 1e-9) {
    const [m, e] = n.toExponential(5).split("e")
    return `${m.replace(/\.?0+$/, "")}E${Number(e) < 0 ? "-" : "+"}${String(Math.abs(Number(e))).padStart(2, "0")}`
  }
  const digits = Math.max(0, 10 - Math.max(1, Math.floor(Math.log10(abs)) + 1))
  return String(Number(n.toFixed(Math.min(digits, 20))))
}

/** True when a format code shows dates or times (used to pick the default format of formula results). */
export function isDateFormat(code: string): boolean {
  return parse(code).some(isDateSection)
}
