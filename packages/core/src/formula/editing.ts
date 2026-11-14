import { MAX_COLS, MAX_ROWS, type Range } from "../address"
import type { RefNode, RefPoint } from "./ast"
import { normalizeFormulaText, tokenize, type Token } from "./lexer"
import { parseRefText } from "./parser"
import { refText } from "./print"
import { FUNCTION_INFO, type FunctionInfo } from "./signatures"

/**
 * What the formula bar and the in-cell editor need while a formula is typed: the references to colour (and outline
 * on the grid), whether a click or an arrow key would point at a cell, F4, and the function the caret is in.
 * Offsets are into the full text, "=" included.
 */

/** Excel cycles through this many reference colours. */
export const REFERENCE_COLORS = 8

export type EditReference = {
  start: number
  end: number
  /** As written: null for the formula's own sheet. */
  sheet: string | null
  range: Range
  /** 0 … REFERENCE_COLORS-1; the same reference written twice gets the same colour. */
  color: number
}

function formulaTokens(text: string): Token[] | null {
  if (!text.startsWith("=")) return null
  const tokens = tokenize(text.slice(1))
  for (const t of tokens) t.start += 1
  return tokens
}

function rangeOfRef(ref: RefNode): Range {
  const to = ref.to ?? ref.from
  const rows = ref.from.row === null ? [0, MAX_ROWS - 1] : [ref.from.row, to.row!]
  const cols = ref.from.col === null ? [0, MAX_COLS - 1] : [ref.from.col, to.col!]
  return {
    r0: Math.min(rows[0], rows[1]),
    r1: Math.max(rows[0], rows[1]),
    c0: Math.min(cols[0], cols[1]),
    c1: Math.max(cols[0], cols[1]),
  }
}

/** The references in a formula being typed, with their colours. 3-D references are left out (no single sheet). */
export function editReferences(text: string): EditReference[] {
  const tokens = formulaTokens(text)
  if (!tokens) return []
  const colors = new Map<string, number>()
  const out: EditReference[] = []
  for (const t of tokens) {
    if (t.kind !== "ref") continue
    const ref = parseRefText(t.text)
    if (!ref || ref.invalid || ref.sheetTo !== null) continue
    const key = t.text.toUpperCase().replace(/\$/g, "")
    let color = colors.get(key)
    if (color === undefined) {
      color = colors.size % REFERENCE_COLORS
      colors.set(key, color)
    }
    out.push({ start: t.start, end: t.start + t.text.length, sheet: ref.sheet, range: rangeOfRef(ref), color })
  }
  return out
}

/**
 * Whether the caret is where a reference can go: right after "=", an operator, "(" or ",". Excel then lets you
 * point at cells with the mouse or the arrow keys (Point mode).
 */
export function canPointAt(text: string, caret: number): boolean {
  const tokens = formulaTokens(text)
  if (!tokens || caret < 1) return false
  let previous: Token | null = null
  for (const t of tokens) {
    if (t.start >= caret) break
    if (t.start + t.text.length > caret) return false // inside a token
    if (t.kind !== "space") previous = t
  }
  if (!previous) return true // right after "="
  if (previous.kind === "op") return previous.text !== "%" && previous.text !== "#"
  return previous.kind === "open" || previous.kind === "comma"
}

const CYCLE: [boolean, boolean][] = [
  [true, true], // $A$1
  [false, true], // A$1
  [true, false], // $A1
  [false, false], // A1
]

function cyclePoint(p: RefPoint, next: [boolean, boolean]): RefPoint {
  return { ...p, colAbs: p.col === null ? false : next[0], rowAbs: p.row === null ? false : next[1] }
}

/**
 * F4: the reference at (or just before) the caret cycles A1 → $A$1 → A$1 → $A1 → A1. Returns the new text and the
 * new reference's span (Excel selects it), or null when the caret is not on a reference.
 */
export function toggleAbsolute(text: string, caret: number): { text: string; start: number; end: number } | null {
  const tokens = formulaTokens(text)
  if (!tokens) return null
  const t = tokens.find((x) => x.kind === "ref" && caret >= x.start && caret <= x.start + x.text.length)
  if (!t) return null
  const ref = parseRefText(t.text)
  if (!ref || ref.invalid) return null
  const state = (p: RefPoint) => [p.col === null ? false : p.colAbs, p.row === null ? false : p.rowAbs]
  const [ca, ra] = state(ref.from)
  const index = CYCLE.findIndex(([c, r]) => c === ca && r === ra)
  let next = CYCLE[(index + 1) % CYCLE.length]
  // Whole columns and rows have one part to fix: A:A ↔ $A:$A.
  if (ref.from.row === null || ref.from.col === null) next = ca || ra ? [false, false] : [true, true]
  const updated: RefNode = {
    ...ref,
    from: cyclePoint(ref.from, next),
    to: ref.to ? cyclePoint(ref.to, next) : null,
  }
  const bang = t.text.lastIndexOf("!")
  const prefix = bang >= 0 ? t.text.slice(0, bang + 1) : ""
  const replacement = prefix + refText({ ...updated, sheet: null, sheetTo: null })
  return {
    text: text.slice(0, t.start) + replacement + text.slice(t.start + t.text.length),
    start: t.start,
    end: t.start + replacement.length,
  }
}

/** The innermost function call around the caret and which argument (0-based) the caret is in. */
export function functionAtCaret(text: string, caret: number): { info: FunctionInfo; arg: number } | null {
  const tokens = formulaTokens(text)
  if (!tokens) return null
  const stack: { name: string | null; arg: number }[] = []
  for (const t of tokens) {
    if (t.start >= caret) break
    if (t.kind === "func") {
      stack.push({ name: t.text.toUpperCase().replace(/^_XLFN\./, ""), arg: 0 })
    } else if (t.kind === "open" || t.kind === "arrayOpen") {
      // "func" tokens hold the name only; their "(" is the next token, not a new group.
      const previous = tokens[tokens.indexOf(t) - 1]
      if (!(t.kind === "open" && previous?.kind === "func")) stack.push({ name: null, arg: 0 })
    } else if (t.kind === "close" || t.kind === "arrayClose") {
      stack.pop()
    } else if (t.kind === "comma" && stack.length) {
      stack[stack.length - 1].arg++
    }
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    const name = stack[i].name
    if (name === null) continue
    const info = FUNCTION_INFO.get(name)
    return info ? { info, arg: stack[i].arg } : null
  }
  return null
}

/** The function name being typed at the caret (for autocomplete): its start and the letters so far. */
export function nameAtCaret(text: string, caret: number): { start: number; prefix: string } | null {
  const tokens = formulaTokens(text)
  if (!tokens) return null
  const t = tokens.find((x) => x.start < caret && x.start + x.text.length === caret)
  if (!t || (t.kind !== "name" && t.kind !== "ref" && t.kind !== "bool")) return null
  if (t.text.includes("!") || t.text.includes(":") || t.text.includes("$")) return null
  if (!/^[A-Za-z_][\w.]*$/.test(t.text)) return null
  return { start: t.start, prefix: t.text }
}

/** Functions whose names start with `prefix` (case-insensitive), for the autocomplete list. */
export function functionsStartingWith(prefix: string, limit = 12): FunctionInfo[] {
  const upper = prefix.toUpperCase()
  const out: FunctionInfo[] = []
  for (const info of FUNCTION_INFO.values()) {
    if (info.name.startsWith(upper)) out.push(info)
    if (out.length >= limit) break
  }
  return out
}

/** A reference to `range` as Point mode writes it (relative; a sheet prefix when pointing at another sheet). */
export function pointedText(range: Range, sheetPrefix = ""): string {
  const point = (row: number, col: number): RefPoint => ({ row, col, rowAbs: false, colAbs: false })
  const single = range.r0 === range.r1 && range.c0 === range.c1
  const node: RefNode = {
    type: "ref",
    sheet: null,
    sheetTo: null,
    from: point(range.r0, range.c0),
    to: single ? null : point(range.r1, range.c1),
  }
  if (range.r0 === 0 && range.r1 === MAX_ROWS - 1) {
    node.from = { row: null, col: range.c0, rowAbs: false, colAbs: false }
    node.to = { row: null, col: range.c1, rowAbs: false, colAbs: false }
  } else if (range.c0 === 0 && range.c1 === MAX_COLS - 1) {
    node.from = { row: range.r0, col: null, rowAbs: false, colAbs: false }
    node.to = { row: range.r1, col: null, rowAbs: false, colAbs: false }
  }
  return sheetPrefix + refText(node)
}

/**
 * What Excel does to a formula when it is entered: function names, references and TRUE/FALSE in capitals, and
 * the closing parentheses that are missing added at the end (=sum(a1:a3 becomes =SUM(A1:A3)).
 */
export function completeFormula(text: string): string {
  if (!text.startsWith("=") || text.length < 2) return text
  const tokens = tokenize(text.slice(1))
  let depth = 0
  for (const t of tokens) {
    if (t.kind === "open") depth++
    else if (t.kind === "close") depth--
  }
  return "=" + normalizeFormulaText(tokens) + (depth > 0 ? ")".repeat(depth) : "")
}
