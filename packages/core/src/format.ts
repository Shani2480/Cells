import { isError, type CellValue } from "./values"

/**
 * Display text for a value. Phase 1 covers General and the formats Excel sets on entry ("0%", "0.00%", "#,##0",
 * "#,##0.00"); the full number-format language replaces this module in Phase 3.
 *
 * `maxChars` is how many digits fit the column: General shortens decimals to fit, as Excel does, and returns null
 * when even the shortest form does not fit (the caller then draws ####).
 */
export function formatValue(value: CellValue | undefined, numberFormat?: string, maxChars = 11): string | null {
  if (value === undefined) {
    return ""
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE"
  }
  if (isError(value)) {
    return value.error
  }
  if (typeof value === "string") {
    return value
  }
  if (numberFormat && numberFormat !== "General") {
    return formatSimple(value, numberFormat)
  }
  return formatGeneral(value, maxChars)
}

/**
 * Excel's General format for a column of `maxChars` characters (11 in a default-width column): integers as they are,
 * decimals rounded to fit, scientific notation for very large or very small numbers.
 */
export function formatGeneral(value: number, maxChars = 11): string | null {
  if (value === 0) {
    return "0"
  }
  const negative = value < 0
  const abs = Math.abs(value)
  const room = maxChars - (negative ? 1 : 0)
  if (room < 1) {
    return null
  }
  const exponent = Math.floor(Math.log10(abs))
  const intDigits = Math.max(1, exponent + 1)
  const useScientific = abs >= 1e11 || exponent < -9 || (abs < 1e-4 && sigDigits(abs) + -exponent + 1 > room)
  if (!useScientific && intDigits <= room) {
    const decimals = Math.max(0, Math.min(room - intDigits - 1, 30))
    let text = abs.toFixed(decimals)
    if (text.includes(".")) {
      text = text.replace(/0+$/, "").replace(/\.$/, "")
    }
    // Rounding can carry into a new digit ("9.99999999999" → "10.0000000"), which still fits after trimming.
    if (text.length <= room && Number(text) !== 0) {
      return (negative ? "-" : "") + text
    }
  }
  // Scientific: mantissa digits to fit, "1.23457E+11".
  for (let decimals = Math.max(0, Math.min(room - 6, 14)); decimals >= 0; decimals--) {
    const [mantissa, exp] = abs.toExponential(decimals).split("e")
    const trimmed = mantissa.includes(".") ? mantissa.replace(/0+$/, "").replace(/\.$/, "") : mantissa
    const power = Number(exp)
    const text = `${trimmed}E${power < 0 ? "-" : "+"}${String(Math.abs(power)).padStart(2, "0")}`
    if (text.length <= room) {
      return (negative ? "-" : "") + text
    }
  }
  return null
}

function sigDigits(abs: number) {
  return abs.toPrecision(15).replace(/e.*$/, "").replace(".", "").replace(/^0+/, "").replace(/0+$/, "").length
}

function formatSimple(value: number, code: string): string {
  const percent = code.endsWith("%")
  const body = percent ? code.slice(0, -1) : code
  const decimals = body.includes(".") ? body.split(".")[1].length : 0
  const grouped = body.includes(",")
  const scaled = percent ? value * 100 : value
  const text = Math.abs(scaled).toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    useGrouping: grouped,
  })
  const negative = scaled < 0 && Number(text.replace(/,/g, "")) !== 0
  return `${negative ? "-" : ""}${text}${percent ? "%" : ""}`
}

export type HorizontalAlign = "left" | "center" | "right"

/** General alignment: numbers right, booleans and errors centred, text left. */
export function generalAlign(value: CellValue | undefined): HorizontalAlign {
  if (typeof value === "number") return "right"
  if (typeof value === "boolean" || isError(value)) return "center"
  return "left"
}
