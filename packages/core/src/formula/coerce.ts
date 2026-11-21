import { errorValue, type CellError } from "../values"
import { parseDateTimeText } from "./dates"
import type { Scalar } from "./types"

export const VALUE = errorValue("#VALUE!")
export const NA = errorValue("#N/A")
export const DIV0 = errorValue("#DIV/0!")
export const NUM = errorValue("#NUM!")
export const REF = errorValue("#REF!")
export const NAME = errorValue("#NAME?")
export const NULL = errorValue("#NULL!")
export const SPILL = errorValue("#SPILL!")
export const CALC = errorValue("#CALC!")

/** The current year, for dates typed without one. The calc context sets it from the clock. */
let currentYear = new Date().getFullYear()
export function setCurrentYear(year: number) {
  currentYear = year
}

/**
 * Text as a number, as Excel converts it in arithmetic and VALUE (en-US): " 12 ", "1,234.5", "$1,000", "(5)",
 * "10%", "1e3", dates, times and date-times. Null when the text is not a number.
 */
export function textToNumber(text: string): number | null {
  let t = text.trim()
  if (t === "") return null
  let negative = false
  if (t.startsWith("(") && t.endsWith(")")) {
    negative = true
    t = t.slice(1, -1).trim()
  }
  let sign = 1
  if (t.startsWith("-") || t.startsWith("+")) {
    if (t[0] === "-") sign = -1
    t = t.slice(1).trim()
  }
  if (t.startsWith("$")) t = t.slice(1).trim()
  if (!negative && sign === 1 && (t.startsWith("-") || t.startsWith("+"))) {
    if (t[0] === "-") sign = -1
    t = t.slice(1).trim()
  }
  let percent = 0
  while (t.endsWith("%")) {
    percent++
    t = t.slice(0, -1).trim()
  }
  let n: number | null = null
  if (/^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(t)) n = Number(t)
  else if (/^\d{1,3}(?:,\d{3})+(?:\.\d*)?$/.test(t)) n = Number(t.replace(/,/g, ""))
  else if (/^\d+ \d+\/\d+$/.test(t)) {
    const [whole, frac] = t.split(" ")
    const [a, b] = frac.split("/").map(Number)
    if (b !== 0 && a < b) n = Number(whole) + a / b
  } else if (percent === 0) n = parseDateTimeText(t, currentYear)
  if (n === null || !Number.isFinite(n)) return null
  n *= sign * (negative ? -1 : 1)
  return percent ? n / 100 ** percent : n
}

/** A scalar as a number for arithmetic: blank is 0, TRUE is 1, numeric text converts, other text is #VALUE!. */
export function toNumber(value: Scalar): number | CellError {
  if (typeof value === "number") return value
  if (value === null) return 0
  if (typeof value === "boolean") return value ? 1 : 0
  if (typeof value === "string") return textToNumber(value) ?? VALUE
  return value
}

/**
 * A number as & and text conversion write it (checked in Excel): 15 significant digits, written out in full while
 * that takes at most 20 characters (0.000000000000000001, 90000000000000000000), else scientific (1E-19, 1E+20).
 */
/** General text for a number: 15 significant digits, in E notation once it would take more than `width` characters. */
export function numberToText(n: number, width = 20): string {
  if (n === 0 || !Number.isFinite(n)) return "0"
  const [mantissa, expText] = Math.abs(n).toPrecision(15).split(/e/i)
  let exp = expText === undefined ? 0 : Number(expText)
  let digits = mantissa.replace(".", "")
  // toPrecision without an exponent: find the exponent from the position of the decimal point.
  if (expText === undefined) {
    const point = mantissa.indexOf(".")
    const intLen = point < 0 ? mantissa.length : point
    exp = intLen - 1
    if (mantissa.startsWith("0")) {
      const firstSig = digits.search(/[1-9]/)
      exp = -firstSig
      digits = digits.slice(firstSig)
    }
  }
  digits = digits.replace(/0+$/, "") || "0"
  let text: string
  if (exp >= 0) {
    text =
      digits.length <= exp + 1
        ? digits + "0".repeat(exp + 1 - digits.length)
        : `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`
  } else {
    text = `0.${"0".repeat(-exp - 1)}${digits}`
  }
  if (text.length > width) {
    const m = digits.length > 1 ? `${digits[0]}.${digits.slice(1)}` : digits
    text = `${m}E${exp < 0 ? "-" : "+"}${String(Math.abs(exp)).padStart(2, "0")}`
  }
  return n < 0 ? `-${text}` : text
}

/** A scalar as text: blank is "", TRUE is "TRUE", numbers in General. */
export function toText(value: Scalar): string | CellError {
  if (typeof value === "string") return value
  if (value === null) return ""
  if (typeof value === "number") return numberToText(value)
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE"
  return value
}

/** A scalar as a logical: numbers are TRUE unless 0, "TRUE"/"FALSE" text converts, other text is #VALUE!. */
export function toBool(value: Scalar): boolean | CellError {
  if (typeof value === "boolean") return value
  if (value === null) return false
  if (typeof value === "number") return value !== 0
  if (typeof value === "string") {
    const upper = value.trim().toUpperCase()
    if (upper === "TRUE") return true
    if (upper === "FALSE") return false
    return VALUE
  }
  return value
}

/** Excel compares numbers to 15 significant digits: =0.1+0.2=0.3 is TRUE. */
export function numbersEqual(a: number, b: number): boolean {
  if (a === b) return true
  if (a === 0 || b === 0) return Math.abs(a - b) < 1e-300
  return Number(a.toPrecision(15)) === Number(b.toPrecision(15))
}

/** The kind order Excel sorts and compares by: numbers < text < logicals; errors are not compared. */
function rank(value: Exclude<Scalar, CellError | null>): number {
  return typeof value === "number" ? 0 : typeof value === "string" ? 1 : 2
}

/**
 * Compares two scalars as the = < > operators do: a blank takes the other side's kind (0, "" or FALSE), text is
 * case-insensitive, different kinds order numbers < text < logicals. Returns -1, 0 or 1.
 */
export function compareScalars(a: Exclude<Scalar, CellError>, b: Exclude<Scalar, CellError>): number {
  if (a === null && b === null) return 0
  if (a === null) a = blankLike(b!)
  if (b === null) b = blankLike(a)
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra < rb ? -1 : 1
  if (typeof a === "number" && typeof b === "number") return numbersEqual(a, b) ? 0 : a < b ? -1 : 1
  if (typeof a === "string" && typeof b === "string") return compareText(a, b)
  return a === b ? 0 : a ? 1 : -1
}

function blankLike(other: Exclude<Scalar, CellError | null>): Exclude<Scalar, CellError | null> {
  return typeof other === "number" ? 0 : typeof other === "string" ? "" : false
}

const collator = new Intl.Collator("en", { sensitivity: "accent" })

/** Text order as Excel compares it: case-insensitive. */
export function compareText(a: string, b: string): number {
  const la = a.toLowerCase()
  const lb = b.toLowerCase()
  if (la === lb) return 0
  const c = collator.compare(la, lb)
  return c === 0 ? (la < lb ? -1 : 1) : c < 0 ? -1 : 1
}
