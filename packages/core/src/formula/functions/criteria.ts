import { isError } from "../../values"
import { compareScalars, compareText, textToNumber } from "../coerce"
import type { Scalar } from "../types"

/**
 * The criteria of COUNTIF, SUMIF, AVERAGEIF and the *IFS functions, shared by all of them: a value (equality, text
 * with wildcards * ? and ~), or an operator before it (">=10", "<>", "=", "<>a*"). Numbers match numbers and
 * numeric text in the criteria ("10" matches 10); text comparison ignores case.
 */
export function compileCriteria(criteria: Scalar): (value: Scalar) => boolean {
  if (isError(criteria)) return (v) => isError(v) && v.error === criteria.error
  if (typeof criteria === "number") return numberEquals(criteria)
  if (typeof criteria === "boolean") return (v) => v === criteria
  if (criteria === null) return (v) => v === 0 // an empty criteria cell matches 0 (Excel treats it as =0)
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(criteria)!
  const op = m[1] ?? "="
  const operand = m[2]
  const operandNumber = textToNumber(operand)
  const upper = operand.toUpperCase()
  const operandBool = upper === "TRUE" ? true : upper === "FALSE" ? false : null
  const errorOperand = /^#/.test(operand) ? upper : null

  if (op === "=" || op === "<>") {
    const negate = op === "<>"
    let test: (v: Scalar) => boolean
    if (operand === "") {
      // "=" matches blank cells only; "<>" matches every non-blank cell; a bare "" matches blank and empty text.
      test = m[1] ? (v) => v === null : (v) => v === null || v === ""
    } else if (operandNumber !== null) {
      test = numberEquals(operandNumber)
    } else if (operandBool !== null) {
      test = (v) => v === operandBool
    } else if (errorOperand) {
      test = (v) => isError(v) && v.error === errorOperand
    } else {
      const pattern = wildcardPattern(operand)
      test = pattern
        ? (v) => typeof v === "string" && pattern.test(v)
        : (v) => typeof v === "string" && compareText(v, operand) === 0
    }
    return negate ? (v) => !test(v) : test
  }
  // < > <= >=: numbers against a number, text against text; blanks never match.
  const holds = (c: number) => (op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0)
  if (operandNumber !== null) return (v) => typeof v === "number" && holds(compareScalars(v, operandNumber))
  if (operandBool !== null) return (v) => typeof v === "boolean" && holds(compareScalars(v, operandBool))
  return (v) => typeof v === "string" && holds(compareText(v, operand))
}

/** Equality with a number also matches text that reads as that number (COUNTIF(A:A,10) counts "10"). */
function numberEquals(n: number) {
  return (v: Scalar) => {
    if (typeof v === "number") return compareScalars(v, n) === 0
    if (typeof v !== "string") return false
    const t = textToNumber(v)
    return t !== null && compareScalars(t, n) === 0
  }
}

/** A case-insensitive regex for text with * and ? (and ~ to escape them), or null when there are none. */
export function wildcardPattern(text: string): RegExp | null {
  if (!/[*?~]/.test(text)) return null
  let source = ""
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch === "~" && i + 1 < text.length && "*?~".includes(text[i + 1])) {
      source += escape(text[++i])
    } else if (ch === "*") source += "[\\s\\S]*"
    else if (ch === "?") source += "[\\s\\S]"
    else source += escape(ch)
  }
  return new RegExp(`^${source}$`, "i")
}

function escape(ch: string) {
  return ch.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&")
}
