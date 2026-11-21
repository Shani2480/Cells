import type { ErrorCode } from "../values"

/** One end of a reference. `abs` marks a `$`. A whole column has no row; a whole row has no column. */
export type RefPoint = { row: number | null; col: number | null; rowAbs: boolean; colAbs: boolean }

/** A reference as written: a cell, an area, whole columns or whole rows, on a sheet (or a sheet span for 3-D). */
export type RefNode = {
  type: "ref"
  sheet: string | null
  /** The last sheet of a 3-D reference (Sheet1:Sheet3!A1). */
  sheetTo: string | null
  from: RefPoint
  to: RefPoint | null
  /** Sheet1!#REF! or a reference that a deletion destroyed. */
  invalid?: boolean
}

export type Node =
  | { type: "number"; value: number }
  | { type: "string"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "error"; code: ErrorCode }
  | RefNode
  | { type: "name"; sheet: string | null; name: string }
  | { type: "func"; name: string; args: Node[] }
  /** A call on what an expression gives (a LAMBDA): =LAMBDA(x,x+1)(5), =LET(f,LAMBDA(…),f)(2). */
  | { type: "call"; callee: Node; args: Node[] }
  | { type: "unary"; op: "-" | "+" | "@"; arg: Node }
  | { type: "percent"; arg: Node }
  | { type: "binary"; op: BinaryOp; left: Node; right: Node }
  | { type: "array"; rows: Node[][] }
  | { type: "spill"; ref: Node }
  | { type: "structured"; text: string }
  | { type: "missing" } // an omitted argument: =IF(A1,,2)
  | { type: "paren"; arg: Node } // kept: Excel's clean-up of tiny sums skips parenthesised expressions

export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=" | ":" | " " | ","
