/**
 * Tokens of an Excel formula. Every character of the input belongs to exactly one token (whitespace included), so
 * joining the tokens' text gives the formula back: rewriting references (insert rows, copy) only swaps the text of
 * reference tokens and keeps the user's spacing, as Excel does.
 */
export type TokenKind =
  | "number"
  | "string"
  | "bool"
  | "error"
  | "ref" // A1, $A$1, A1:B2, A:A, 3:5, with an optional sheet prefix, also R1C1 forms when asked for
  | "name" // a defined name or a bare word (resolved by the parser/evaluator)
  | "func" // a name followed by "("
  | "op" // + - * / ^ & = <> < > <= >= % : @
  | "comma"
  | "semicolon" // row separator inside array constants
  | "open"
  | "close"
  | "arrayOpen"
  | "arrayClose"
  | "space"
  | "spill" // the # after a reference (A1#)
  | "structured" // Table1[Column] (parsed in Phase 4)
  | "invalid"

export type Token = { kind: TokenKind; text: string; start: number }

const ERRORS = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
]

/** A sheet prefix: Sheet1!, 'My Sheet'!, Sheet1:Sheet3!, '[Book]Sheet'! */
const SHEET = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\u00a0-\uffff][\w.\u00a0-\uffff]*(?::[A-Za-z_\u00a0-\uffff][\w.\u00a0-\uffff]*)?)!`
const COL = String.raw`\$?[A-Za-z]{1,3}`
const ROW = String.raw`\$?[0-9]{1,7}`
const CELL = `${COL}${ROW}`
// Longest forms first: an area, a whole-column span, a whole-row span, a single cell.
const REF_BODY = `(?:${CELL}(?::${CELL})?|${COL}:${COL}|${ROW}:${ROW})`
const REF = new RegExp(`^(?:${SHEET})?${REF_BODY}(?![\\w(.!])`, "y")
const SHEET_ONLY = new RegExp(`^${SHEET}`)
const R1C1_PART = String.raw`(?:R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?|R(?:\[-?\d+\]|\d+)?|C(?:\[-?\d+\]|\d+)?)`
const R1C1 = new RegExp(`^(?:${SHEET})?${R1C1_PART}(?::${R1C1_PART})?(?![\\w(.!])`, "iy")
const NUMBER = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/
const NAME = /^[A-Za-z_\\\u00a0-\uffff][\w.?\\\u00a0-\uffff]*/

export type LexOptions = { r1c1?: boolean }

/** Splits a formula (without its leading "=") into tokens. Never throws: unknown characters become "invalid". */
export function tokenize(source: string, options: LexOptions = {}): Token[] {
  const tokens: Token[] = []
  let i = 0
  let arrayDepth = 0
  const push = (kind: TokenKind, text: string) => {
    tokens.push({ kind, text, start: i })
    i += text.length
  }
  while (i < source.length) {
    const ch = source[i]
    const rest = source.slice(i)
    if (ch === " " || ch === "\n" || ch === "\r" || ch === "\t") {
      push("space", /^[ \n\r\t]+/.exec(rest)![0])
      continue
    }
    if (ch === '"') {
      let j = i + 1
      while (j < source.length) {
        if (source[j] === '"') {
          if (source[j + 1] === '"') j += 2
          else break
        } else j++
      }
      push("string", source.slice(i, Math.min(j + 1, source.length)))
      continue
    }
    if (ch === "#") {
      const prev = tokens.at(-1)
      if (prev && (prev.kind === "ref" || prev.kind === "close" || prev.kind === "name")) {
        push("spill", "#")
        continue
      }
      const upper = rest.toUpperCase()
      const error = ERRORS.find((e) => upper.startsWith(e))
      // A reference whose sheet was deleted: #REF!B2, #REF!A1:C3 (as Excel writes it). One ref token, invalid.
      if (error === "#REF!") {
        const body =
          /^(?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w.(!])/.exec(
            rest.slice(5),
          )
        if (body) {
          push("ref", rest.slice(0, 5 + body[0].length))
          continue
        }
      }
      push(error ? "error" : "invalid", error ? rest.slice(0, error.length) : ch)
      continue
    }
    if (ch === "{") {
      arrayDepth++
      push("arrayOpen", ch)
      continue
    }
    if (ch === "}") {
      arrayDepth = Math.max(0, arrayDepth - 1)
      push("arrayClose", ch)
      continue
    }
    if (ch === "(") {
      push("open", ch)
      continue
    }
    if (ch === ")") {
      push("close", ch)
      continue
    }
    if (ch === ",") {
      push("comma", ch)
      continue
    }
    if (ch === ";") {
      push(arrayDepth > 0 ? "semicolon" : "invalid", ch)
      continue
    }
    const two = source.slice(i, i + 2)
    if (two === "<>" || two === "<=" || two === ">=") {
      push("op", two)
      continue
    }
    if ("+-*/^&=<>%@:".includes(ch)) {
      push("op", ch)
      continue
    }
    // References before numbers (3:5 is a row span) and before names (A1 is not a name).
    const refMatch = matchAt(options.r1c1 ? R1C1 : REF, source, i) ?? (options.r1c1 ? null : matchAt(REF, source, i))
    if (refMatch && !isBoolWord(refMatch)) {
      push("ref", refMatch)
      continue
    }
    if (ch === "'" || /[A-Za-z_\u00a0-\uffff]/.test(ch)) {
      // A sheet-qualified name or a #REF! after a sheet: Sheet1!#REF!, 'x'!Name
      const sheet = SHEET_ONLY.exec(rest)?.[0]
      if (sheet) {
        const after = rest.slice(sheet.length)
        if (after.toUpperCase().startsWith("#REF!")) {
          push("ref", sheet + after.slice(0, 5))
          continue
        }
        const name = NAME.exec(after)?.[0]
        if (name) {
          push("name", sheet + name)
          continue
        }
      }
    }
    const number = NUMBER.exec(rest)?.[0]
    if (number) {
      push("number", number)
      continue
    }
    const name = NAME.exec(rest)?.[0]
    if (name) {
      // Table1[Column] / Table1[[#Headers],[A]] / [@Col]
      if (source[i + name.length] === "[") {
        push("structured", name + bracketed(source, i + name.length))
        continue
      }
      const upper = name.toUpperCase()
      let j = i + name.length
      while (source[j] === " ") j++
      if (source[j] === "(" && source[i + name.length] === "(") {
        push("func", name)
        continue
      }
      if (upper === "TRUE" || upper === "FALSE") {
        push("bool", name)
        continue
      }
      push("name", name)
      continue
    }
    if (ch === "[") {
      push("structured", bracketed(source, i))
      continue
    }
    push("invalid", ch)
  }
  return tokens
}

function isBoolWord(text: string) {
  const upper = text.toUpperCase()
  return upper === "TRUE" || upper === "FALSE"
}

function matchAt(pattern: RegExp, source: string, index: number): string | null {
  pattern.lastIndex = 0
  const match = pattern.exec(source.slice(index))
  return match ? match[0] : null
}

function bracketed(source: string, from: number): string {
  let depth = 0
  let j = from
  for (; j < source.length; j++) {
    if (source[j] === "'" && j + 1 < source.length) {
      j++
      continue
    }
    if (source[j] === "[") depth++
    else if (source[j] === "]") {
      depth--
      if (depth === 0) return source.slice(from, j + 1)
    }
  }
  return source.slice(from)
}

/** The formula with references, function names, booleans and errors in upper case, as Excel stores it. */
export function normalizeFormulaText(tokens: Token[]): string {
  return tokens
    .map((t) => {
      if (t.kind === "func" || t.kind === "bool" || t.kind === "error") return t.text.toUpperCase()
      if (t.kind === "ref") return upperRef(t.text)
      return t.text
    })
    .join("")
}

/** Upper-cases a reference but leaves a quoted sheet name alone. */
function upperRef(text: string) {
  const bang = text.lastIndexOf("!")
  if (bang < 0) return text.toUpperCase()
  return text.slice(0, bang + 1) + text.slice(bang + 1).toUpperCase()
}
