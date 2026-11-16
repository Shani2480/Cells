import { MAX_COLS, MAX_ROWS, columnIndex } from "../address"
import { ERROR_CODES, truncateDigits, type ErrorCode } from "../values"
import type { BinaryOp, Node, RefNode, RefPoint } from "./ast"
import { tokenize, type Token } from "./lexer"

export type ParseResult = { ok: true; ast: Node; tokens: Token[] } | { ok: false; message: string; position: number }

class ParseError extends Error {
  constructor(
    message: string,
    readonly position: number,
  ) {
    super(message)
  }
}

/** Parses a formula without its leading "=". */
export function parseFormula(source: string): ParseResult {
  const tokens = tokenize(source)
  try {
    const parser = new Parser(tokens, source.length)
    const ast = parser.parseTop()
    return { ok: true, ast, tokens }
  } catch (error) {
    if (error instanceof ParseError) return { ok: false, message: error.message, position: error.position }
    throw error
  }
}

/** Tokens that can end an operand, and tokens that can start one: a space between them is the intersection. */
const ENDS_OPERAND = new Set(["ref", "name", "close", "spill", "structured"])
const STARTS_OPERAND = new Set(["ref", "name", "open", "func", "structured"])

class Parser {
  private tokens: Token[]
  private i = 0

  constructor(
    all: Token[],
    private readonly length: number,
  ) {
    // Keep a space only where it is an operator (between two reference operands); the rest is layout.
    this.tokens = []
    for (let k = 0; k < all.length; k++) {
      const t = all[k]
      if (t.kind !== "space") {
        this.tokens.push(t)
        continue
      }
      const prev = this.tokens[this.tokens.length - 1]
      const next = all[k + 1]
      if (prev && next && ENDS_OPERAND.has(prev.kind) && STARTS_OPERAND.has(next.kind)) this.tokens.push(t)
    }
  }

  private peek(): Token | undefined {
    return this.tokens[this.i]
  }

  private next(): Token {
    const t = this.tokens[this.i++]
    if (!t) throw new ParseError("The formula ends too early.", this.length)
    return t
  }

  private isOp(text: string) {
    const t = this.peek()
    return t?.kind === "op" && t.text === text
  }

  private fail(message: string): never {
    throw new ParseError(message, this.peek()?.start ?? this.length)
  }

  parseTop(): Node {
    if (this.tokens.length === 0) this.fail("The formula is empty.")
    const node = this.comparison()
    if (this.i < this.tokens.length) {
      const t = this.peek()!
      this.fail(t.kind === "close" ? "There is an extra closing parenthesis." : `Unexpected "${t.text}".`)
    }
    return node
  }

  private binaryLevel(ops: string[], operand: () => Node): Node {
    let left = operand()
    for (;;) {
      const t = this.peek()
      if (t?.kind !== "op" || !ops.includes(t.text)) return left
      this.i++
      const right = operand()
      left = { type: "binary", op: t.text as BinaryOp, left, right }
    }
  }

  private comparison(): Node {
    return this.binaryLevel(["=", "<>", "<", ">", "<=", ">="], () => this.concat())
  }

  private concat(): Node {
    return this.binaryLevel(["&"], () => this.additive())
  }

  private additive(): Node {
    return this.binaryLevel(["+", "-"], () => this.multiplicative())
  }

  private multiplicative(): Node {
    return this.binaryLevel(["*", "/"], () => this.power())
  }

  /** Excel evaluates ^ left to right: 2^3^2 is 64. */
  private power(): Node {
    return this.binaryLevel(["^"], () => this.percent())
  }

  private percent(): Node {
    let node = this.unary()
    while (this.isOp("%")) {
      this.i++
      node = { type: "percent", arg: node }
    }
    return node
  }

  /** Negation binds tighter than ^ and %: -2^2 is 4. */
  private unary(): Node {
    const t = this.peek()
    if (t?.kind === "op" && (t.text === "-" || t.text === "+" || t.text === "@")) {
      this.i++
      return { type: "unary", op: t.text, arg: this.unary() }
    }
    return this.intersection()
  }

  private intersection(): Node {
    let left = this.range()
    while (this.peek()?.kind === "space") {
      this.i++
      const right = this.range()
      left = { type: "binary", op: " ", left, right }
    }
    return left
  }

  private range(): Node {
    let left = this.postfix()
    while (this.isOp(":")) {
      this.i++
      const right = this.postfix()
      left = { type: "binary", op: ":", left, right }
    }
    return left
  }

  private postfix(): Node {
    let node = this.primary()
    while (this.peek()?.kind === "spill") {
      this.i++
      node = { type: "spill", ref: node }
    }
    return node
  }

  private primary(): Node {
    const t = this.next()
    switch (t.kind) {
      case "number":
        return { type: "number", value: Number(truncateDigits(t.text)) }
      case "string":
        if (!/^"(?:[^"]|"")*"$/.test(t.text))
          throw new ParseError("A text value is missing its closing quote.", t.start)
        return { type: "string", value: t.text.slice(1, -1).replace(/""/g, '"') }
      case "bool":
        return { type: "bool", value: t.text.toUpperCase() === "TRUE" }
      case "error":
        return { type: "error", code: t.text.toUpperCase() as ErrorCode }
      case "ref":
        return parseRefText(t.text) ?? { type: "name", sheet: null, name: t.text }
      case "name":
        return nameNode(t.text)
      case "structured":
        return { type: "structured", text: t.text }
      case "func":
        return this.call(t)
      case "open": {
        const first = this.comparison()
        let node = first
        let union = false
        while (this.peek()?.kind === "comma") {
          this.i++
          union = true
          node = { type: "binary", op: ",", left: node, right: this.comparison() }
        }
        if (this.peek()?.kind !== "close") this.fail("A closing parenthesis is missing.")
        this.i++
        return union ? node : { type: "paren", arg: node }
      }
      case "arrayOpen":
        return this.array()
      default:
        throw new ParseError(
          t.kind === "close" ? "There is an extra closing parenthesis." : `Unexpected "${t.text}".`,
          t.start,
        )
    }
  }

  private call(t: Token): Node {
    const name = t.text.toUpperCase().replace(/^_XLFN\.(?:_XLWS\.)?/, "")
    this.next() // "("
    const args: Node[] = []
    if (this.peek()?.kind === "close") {
      this.i++
      return { type: "func", name, args }
    }
    for (;;) {
      const p = this.peek()
      if (p?.kind === "comma" || p?.kind === "close") args.push({ type: "missing" })
      else args.push(this.comparison())
      const sep = this.peek()
      if (sep?.kind === "comma") {
        this.i++
        continue
      }
      if (sep?.kind === "close") {
        this.i++
        break
      }
      this.fail(sep ? `Unexpected "${sep.text}".` : `${name} is missing its closing parenthesis.`)
    }
    if (args.length > 255) this.fail("A function can take at most 255 arguments.")
    return { type: "func", name, args }
  }

  private array(): Node {
    const rows: Node[][] = [[]]
    for (;;) {
      const t = this.next()
      let value: Node
      if (t.kind === "op" && (t.text === "-" || t.text === "+")) {
        const n = this.next()
        if (n.kind !== "number")
          throw new ParseError("Array constants hold only numbers, text, logicals and errors.", n.start)
        value = { type: "number", value: (t.text === "-" ? -1 : 1) * Number(n.text) }
      } else if (t.kind === "number") value = { type: "number", value: Number(truncateDigits(t.text)) }
      else if (t.kind === "string") value = { type: "string", value: t.text.slice(1, -1).replace(/""/g, '"') }
      else if (t.kind === "bool") value = { type: "bool", value: t.text.toUpperCase() === "TRUE" }
      else if (t.kind === "error") value = { type: "error", code: t.text.toUpperCase() as ErrorCode }
      else throw new ParseError("Array constants hold only numbers, text, logicals and errors.", t.start)
      rows[rows.length - 1]!.push(value)
      const sep = this.next()
      if (sep.kind === "comma") continue
      if (sep.kind === "semicolon") {
        rows.push([])
        continue
      }
      if (sep.kind === "arrayClose") break
      throw new ParseError(`Unexpected "${sep.text}" in an array constant.`, sep.start)
    }
    const width = rows[0].length
    if (rows.some((row) => row.length !== width)) this.fail("Every row of an array constant needs the same length.")
    return { type: "array", rows }
  }
}

function nameNode(text: string): Node {
  const bang = text.lastIndexOf("!")
  if (bang < 0) return { type: "name", sheet: null, name: text }
  return { type: "name", sheet: unquoteSheet(text.slice(0, bang)), name: text.slice(bang + 1) }
}

export function unquoteSheet(text: string): string {
  return text.startsWith("'") && text.endsWith("'") ? text.slice(1, -1).replace(/''/g, "'") : text
}

/** Parses the text of a reference token (A1 style). Null when it is out of the sheet (then it is a name). */
export function parseRefText(text: string): RefNode | null {
  let sheet: string | null = null
  let sheetTo: string | null = null
  let body = text
  // The prefix ends at the last "!" before the body; #REF! carries its own.
  const end = /#REF!$/i.test(text) ? text.length - 5 : text.length
  const bang = text.lastIndexOf("!", end - 1)
  if (bang >= 0) {
    const prefix = text.slice(0, bang)
    body = text.slice(bang + 1)
    if (prefix.startsWith("'")) {
      const inner = unquoteSheet(prefix)
      const colon = splitSheetSpan(inner)
      sheet = colon[0]
      sheetTo = colon[1]
    } else {
      const parts = prefix.split(":")
      sheet = parts[0]
      sheetTo = parts[1] ?? null
    }
  }
  if (body.toUpperCase() === "#REF!" || sheet?.toUpperCase() === "#REF") {
    return {
      type: "ref",
      sheet,
      sheetTo,
      from: { row: 0, col: 0, rowAbs: false, colAbs: false },
      to: null,
      invalid: true,
    }
  }
  const [a, b] = body.split(":")
  const from = parsePoint(a)
  const to = b === undefined ? null : parsePoint(b)
  if (!from || (b !== undefined && !to)) return null
  // A single point must be a cell; spans must pair columns with columns and rows with rows.
  if (!to && (from.row === null || from.col === null)) return null
  if (to && (from.row === null) !== (to.row === null)) return null
  if (to && (from.col === null) !== (to.col === null)) return null
  return { type: "ref", sheet, sheetTo, from, to }
}

/** 'Jan:Mar' → ["Jan", "Mar"]; a quoted name may itself contain a colon only in a 3-D span. */
function splitSheetSpan(inner: string): [string, string | null] {
  const colon = inner.indexOf(":")
  return colon < 0 ? [inner, null] : [inner.slice(0, colon), inner.slice(colon + 1)]
}

function parsePoint(text: string): RefPoint | null {
  const m = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/.exec(text)
  if (!m || (!m[2] && !m[4])) return null
  if (m[2] && !m[4] && m[3]) return null // "A$" is not a reference
  const col = m[2] ? columnIndex(m[2].toUpperCase()) : null
  const row = m[4] ? Number(m[4]) - 1 : null
  if (col !== null && (col < 0 || col >= MAX_COLS)) return null
  if (row !== null && (row < 0 || row >= MAX_ROWS)) return null
  if (!m[2]) return { row, col: null, rowAbs: m[1] === "$", colAbs: false }
  return { row, col, colAbs: m[1] === "$", rowAbs: m[3] === "$" }
}

export const ERROR_SET: ReadonlySet<string> = new Set(ERROR_CODES)
