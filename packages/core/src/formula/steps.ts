import type { Node } from "./ast"
import { printFormula } from "./print"
import type { CellError } from "../values"
import { isLambda, isMatrix, isRef, type Scalar, type Value } from "./types"

/**
 * Evaluate Formula (Formulas › Formula Auditing): a formula worked out one part at a time, as Excel's dialog does.
 * Each step evaluates the innermost, leftmost part whose own parts are all values: a cell reference, a name, an
 * operator between two values, a function whose arguments are values or ranges. A range stays as written until the
 * function or operator that uses it is evaluated.
 */

const LITERAL = new Set(["number", "string", "bool", "error", "missing", "array"])

const LAZY = new Set(["IF", "IFERROR", "IFNA", "CHOOSE"])

export const isLiteral = (n: Node) => LITERAL.has(n.type)

/** A reference to more than one cell, or a range built with ":" or a union: kept until something uses it. */
export function isRange(n: Node): boolean {
  if (n.type === "ref") return Boolean(n.to) || n.from.row === null || n.from.col === null || Boolean(n.sheetTo)
  if (n.type === "binary" && (n.op === ":" || n.op === "," || n.op === " "))
    return (n.left.type === "ref" || isRange(n.left)) && (n.right.type === "ref" || isRange(n.right))
  return n.type === "spill"
}

function children(n: Node): Node[] {
  switch (n.type) {
    case "func":
      return n.args
    case "call":
      return [n.callee, ...n.args]
    case "unary":
    case "percent":
    case "paren":
      return [n.arg]
    case "binary":
      return [n.left, n.right]
    default:
      return []
  }
}

/** The part the next step evaluates, or null when the formula is down to a value. */
export function nextPart(root: Node): Node | null {
  const find = (n: Node): Node | null => {
    if (isLiteral(n)) return null
    if (n !== root && isRange(n)) return null
    // IF, IFERROR, IFNA and CHOOSE work out their first argument, then themselves: the branch not taken isn't shown.
    const parts = n.type === "func" && LAZY.has(n.name.toUpperCase()) ? n.args.slice(0, 1) : children(n)
    for (const child of parts) {
      const hit = find(child)
      if (hit) return hit
    }
    return n
  }
  return find(root)
}

const scalarNode = (v: Scalar): Node => {
  if (v === null) return { type: "number", value: 0 }
  if (typeof v === "number") return { type: "number", value: v }
  if (typeof v === "string") return { type: "string", value: v }
  if (typeof v === "boolean") return { type: "bool", value: v }
  return { type: "error", code: (v as CellError).error }
}

/** A calculated value as the formula's text shows it: a value, or an array of values. */
export function valueNode(v: Value): Node {
  if (isRef(v)) return { type: "error", code: "#VALUE!" }
  if (isLambda(v)) return { type: "error", code: "#CALC!" }
  if (isMatrix(v)) {
    if (v.rows === 1 && v.cols === 1) return scalarNode(v.data[0])
    const rows: Node[][] = []
    for (let r = 0; r < v.rows; r++) rows.push(v.data.slice(r * v.cols, (r + 1) * v.cols).map(scalarNode))
    return { type: "array", rows }
  }
  return scalarNode(v)
}

/** The formula with one part replaced by its value. */
export function replacePart(root: Node, part: Node, value: Node): Node {
  if (root === part) return value
  const map = (n: Node): Node => {
    if (n === part) return value
    switch (n.type) {
      case "func":
        return { ...n, args: n.args.map(map) }
      case "call":
        return { ...n, callee: map(n.callee), args: n.args.map(map) }
      case "unary":
      case "percent":
      case "paren":
        return { ...n, arg: map(n.arg) }
      case "binary":
        return { ...n, left: map(n.left), right: map(n.right) }
      default:
        return n
    }
  }
  return map(root)
}

/** The formula's text and where the part to evaluate next sits in it (the dialog underlines it). */
export function stepText(root: Node, part: Node | null): { text: string; start: number; end: number } {
  if (!part) {
    const text = printFormula(root)
    return { text, start: 0, end: 0 }
  }
  const marker: Node = { type: "name", sheet: null, name: "\u0001" }
  const marked = printFormula(replacePart(root, part, marker))
  const start = marked.indexOf("\u0001")
  const inner = printFormula(part)
  return { text: marked.slice(0, start) + inner + marked.slice(start + 1), start, end: start + inner.length }
}

/** Whether an evaluated value is an error (the dialog says where the error came from). */
export const isErrorNode = (n: Node) =>
  n.type === "error" || (n.type === "array" && n.rows.flat().some((x) => x.type === "error"))
