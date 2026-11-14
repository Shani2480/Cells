/** Excel's error values. */
export const ERROR_CODES = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#SPILL!",
  "#CALC!",
  "#GETTING_DATA",
] as const

export type ErrorCode = (typeof ERROR_CODES)[number]

export type CellError = { readonly error: ErrorCode }

/** What a cell holds after parsing or calculation. "Empty" is `undefined` in the store, never a value. */
export type CellValue = number | string | boolean | CellError

const ERRORS = new Map<string, CellError>(ERROR_CODES.map((code) => [code, Object.freeze({ error: code })]))

export function errorValue(code: ErrorCode): CellError {
  return ERRORS.get(code)!
}

export function isError(value: unknown): value is CellError {
  return typeof value === "object" && value !== null && "error" in value
}

/**
 * The result of reading what a person typed into a cell.
 * - `value` is what the cell holds (a formula's value is computed by the calc engine from Phase 2; until then the
 *   formula text stands in for it).
 * - `input` is kept only when it differs from what the value would print as: formulas, a leading apostrophe,
 *   "1,000" typed with separators, "10%". The formula bar shows it.
 * - `numberFormat` is the format Excel applies on entry ("10%" → "0%"); the caller applies it only when the cell
 *   has no format of its own (General).
 */
export type ParsedInput = {
  value: CellValue
  input?: string
  formula?: boolean
  numberFormat?: string
}

const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/
// Thousands separators as Excel (en-US) accepts them: groups of exactly three digits after the first.
const GROUPED_RE = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d*)?$/

/** Parses typed text the way Excel does on entry (en-US); dates and currencies arrive with number formats in Phase 3. */
export function parseInput(text: string): ParsedInput | null {
  if (text === "") {
    return null
  }
  if (text.startsWith("'")) {
    return { value: text.slice(1), input: text }
  }
  if (text.startsWith("=") && text.length > 1) {
    // The value is a placeholder until the engine calculates (at once, or when the worker answers): show nothing.
    return { value: "", input: text, formula: true }
  }
  const trimmed = text.trim()
  const upper = trimmed.toUpperCase()
  if (upper === "TRUE" || upper === "FALSE") {
    return { value: upper === "TRUE" }
  }
  const error = ERRORS.get(upper)
  if (error) {
    return { value: error }
  }
  const number = parseNumberText(trimmed)
  if (number) {
    return number
  }
  return { value: text }
}

/**
 * Numbers never keep their typed text: Excel's formula bar shows "1000" for "1,000" and "10%" for "10%" from the
 * value and the number format set on entry.
 */
function parseNumberText(text: string): ParsedInput | null {
  let body = text
  let percent = false
  if (body.endsWith("%")) {
    percent = true
    body = body.slice(0, -1).trimEnd()
  }
  let grouped = false
  if (GROUPED_RE.test(body)) {
    grouped = true
    body = body.replace(/,/g, "")
  }
  if (!NUMBER_RE.test(body)) {
    return null
  }
  let value = Number(truncateDigits(body))
  if (!Number.isFinite(value)) {
    return null
  }
  if (percent) {
    value = roundTo15(value / 100)
  }
  const parsed: ParsedInput = { value }
  if (percent) {
    const decimals = body.includes(".") ? body.split(".")[1].replace(/[eE].*/, "").length : 0
    parsed.numberFormat = decimals > 0 ? `0.${"0".repeat(Math.min(decimals, 30))}%` : "0%"
  } else if (grouped) {
    parsed.numberFormat = body.includes(".") ? "#,##0.00" : "#,##0"
  }
  return parsed
}

/**
 * Excel keeps the first 15 significant digits of a typed number and drops the rest (it does not round):
 * "1234567890123456789" is stored as 1234567890123450000.
 */
export function truncateDigits(body: string): string {
  const match = /^([+-]?)(\d*)\.?(\d*)(?:[eE]([+-]?\d+))?$/.exec(body)
  if (!match) return body
  const [, sign, intPart, fracPart, exp] = match
  const digits = intPart + fracPart
  const firstSig = digits.search(/[1-9]/)
  if (firstSig === -1 || digits.length - firstSig <= 15) return body
  const kept = digits.slice(0, firstSig + 15) + "0".repeat(digits.length - firstSig - 15)
  const mantissa = `${kept.slice(0, intPart.length)}.${kept.slice(intPart.length)}`
  return `${sign}${mantissa}${exp !== undefined ? `e${exp}` : ""}`
}

/** Rounds to 15 significant digits (after arithmetic such as a typed percentage divided by 100). */
export function roundTo15(value: number): number {
  if (value === 0 || !Number.isFinite(value)) {
    return value
  }
  return Number(value.toPrecision(15))
}

/** The text the formula bar and the in-cell editor start from when a cell has no stored input. */
export function editText(value: CellValue | undefined): string {
  if (value === undefined) {
    return ""
  }
  if (typeof value === "number") {
    return numberToEditText(value)
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE"
  }
  if (isError(value)) {
    return value.error
  }
  return value
}

/** Up to 15 significant digits, no exponent below 1e21 (as Excel's formula bar). */
export function numberToEditText(value: number): string {
  if (value === 0) {
    return "0"
  }
  const precise = Number(value.toPrecision(15))
  const abs = Math.abs(precise)
  if (abs >= 1e21 || abs < 1e-9) {
    return precise.toExponential().toUpperCase()
  }
  return String(precise).includes("e") ? precise.toFixed(20).replace(/\.?0+$/, "") : String(precise)
}

/** Comparison key for "same value" checks (fill series, find). */
export function valueKind(value: CellValue): "number" | "text" | "boolean" | "error" {
  if (typeof value === "number") return "number"
  if (typeof value === "boolean") return "boolean"
  if (isError(value)) return "error"
  return "text"
}
