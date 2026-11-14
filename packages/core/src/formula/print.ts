import { columnName } from "../address"
import type { Node, RefNode, RefPoint } from "./ast"

/** Excel quotes a sheet name unless it is a plain word that does not look like a reference or a number. */
export function quoteSheet(name: string): string {
  const plain = /^[A-Za-z_\u00a0-\uffff][\w.\u00a0-\uffff]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
  const r1c1 = /^(?:R\d*C?\d*|C\d*)$/i.test(name)
  return plain && !r1c1 && !/^(TRUE|FALSE)$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

export function sheetPrefix(sheet: string | null, sheetTo: string | null = null): string {
  if (sheet === null) return ""
  if (sheetTo === null) return `${quoteSheet(sheet)}!`
  const both = `${sheet}:${sheetTo}`
  return quoteSheet(sheet) === sheet && quoteSheet(sheetTo) === sheetTo ? `${both}!` : `'${both.replace(/'/g, "''")}'!`
}

function pointText(p: RefPoint): string {
  const col = p.col === null ? "" : `${p.colAbs ? "$" : ""}${columnName(p.col)}`
  const row = p.row === null ? "" : `${p.rowAbs ? "$" : ""}${p.row + 1}`
  return col + row
}

export function refText(ref: RefNode): string {
  const prefix = sheetPrefix(ref.sheet, ref.sheetTo)
  if (ref.invalid) return `${prefix}#REF!`
  return prefix + pointText(ref.from) + (ref.to ? `:${pointText(ref.to)}` : "")
}

const PRECEDENCE: Record<string, number> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
  ",": 8,
  " ": 9,
  ":": 10,
}

function valueText(node: Node): string {
  switch (node.type) {
    case "number":
      return numberText(node.value)
    case "string":
      return `"${node.value.replace(/"/g, '""')}"`
    case "bool":
      return node.value ? "TRUE" : "FALSE"
    case "error":
      return node.code
    default:
      return printFormula(node)
  }
}

function numberText(n: number): string {
  return String(n).toUpperCase().replace("E+", "E+")
}

/** Prints an AST in Excel's canonical spelling (no layout spaces). Used by tests and for generated formulas. */
export function printFormula(node: Node, parent = 0): string {
  switch (node.type) {
    case "number":
    case "string":
    case "bool":
    case "error":
      return valueText(node)
    case "ref":
      return refText(node)
    case "name":
      return sheetPrefix(node.sheet) + node.name
    case "structured":
      return node.text
    case "missing":
      return ""
    case "paren":
      return `(${printFormula(node.arg)})`
    case "func":
      return `${node.name}(${node.args.map((a) => printFormula(a)).join(",")})`
    case "unary": {
      const text = node.op + printFormula(node.arg, 6)
      return parent > 6 ? `(${text})` : text
    }
    case "percent": {
      const text = printFormula(node.arg, 6) + "%"
      return parent > 7 ? `(${text})` : text
    }
    case "spill":
      return printFormula(node.ref, 11) + "#"
    case "array":
      return `{${node.rows.map((row) => row.map(valueText).join(",")).join(";")}}`
    case "binary": {
      const p = PRECEDENCE[node.op]
      // Left-associative: the right side needs parentheses at the same level.
      const text = printFormula(node.left, p) + node.op + printFormula(node.right, p + 1)
      return parent > p || node.op === "," ? `(${text})` : text
    }
  }
}
