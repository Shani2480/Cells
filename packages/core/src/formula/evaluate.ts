import { MAX_COLS, MAX_ROWS } from "../address"
import { errorValue, isError, type CellError, type ErrorCode } from "../values"
import type { BinaryOp, Node, RefNode } from "./ast"
import { compareScalars, DIV0, NA, NAME, NULL, NUM, REF, toNumber, toText, VALUE } from "./coerce"
import type { EvalContext } from "./context"
import { FUNCTIONS, type FnDef } from "./functions/registry"
import { parseStructured, structuredArea, type StructuredSpec } from "../table/table"
import {
  areaCols,
  areaRows,
  isMatrix,
  isRef,
  matrix,
  type Area,
  type Matrix,
  type RefValue,
  type Scalar,
  type Value,
} from "./types"

/** Evaluates a formula's AST to its result (a scalar, or an array that spills). */
export function evaluateFormula(node: Node, ctx: EvalContext): Value {
  // Excel cleans up cancellation only in the formula's last + or − (not inside parentheses or functions).
  const value =
    node.type === "binary" && (node.op === "+" || node.op === "-")
      ? binary(node.op, node.left, node.right, ctx, true)
      : evaluate(node, ctx)
  return finalValue(value, ctx)
}

/** A formula's result as the cell holds it: a reference becomes its value(s); a 1×1 array a scalar. */
export function finalValue(value: Value, ctx: EvalContext): Value {
  if (isRef(value)) {
    if (value.areas.length > 1) return VALUE
    const area = value.areas[0]
    if (area.r0 === area.r1 && area.c0 === area.c1) return blankToZero(ctx.cell(area.sheet, area.r0, area.c0))
    return mapMatrix(refToMatrix(value, ctx), blankToZero)
  }
  if (isMatrix(value)) {
    if (value.rows === 1 && value.cols === 1) return blankToZero(value.data[0])
    return mapMatrix(value, blankToZero)
  }
  return value
}

/** A formula that points at a blank cell shows 0 (=A1 with A1 empty). */
function blankToZero(value: Scalar): Scalar {
  return value === null ? 0 : value
}

export function evaluate(node: Node, ctx: EvalContext): Value {
  switch (node.type) {
    case "number":
      return node.value
    case "string":
      return node.value
    case "bool":
      return node.value
    case "error":
      return errorValue(node.code)
    case "missing":
      return null
    case "paren":
      return evaluate(node.arg, ctx)
    case "ref":
      return resolveRef(node, ctx)
    case "name":
      return resolveName(node.name, node.sheet, ctx)
    case "structured":
      return resolveStructured(node.text, ctx)
    case "array":
      return matrix(
        node.rows.length,
        node.rows[0].length,
        node.rows.flat().map((n) => evaluate(n, ctx) as Scalar),
      )
    case "func":
      return callFunction(node.name, node.args, ctx)
    case "unary": {
      if (node.op === "@") return implicitIntersection(evaluate(node.arg, ctx), ctx)
      const arg = evaluate(node.arg, ctx)
      if (node.op === "+") return arg
      return lift1(arg, ctx, (v) => {
        const n = toNumber(v)
        return typeof n === "number" ? -n : n
      })
    }
    case "percent":
      return lift1(evaluate(node.arg, ctx), ctx, (v) => {
        const n = toNumber(v)
        return typeof n === "number" ? n / 100 : n
      })
    case "spill": {
      const target = evaluate(node.ref, ctx)
      if (!isRef(target) || target.areas.length !== 1) return REF
      const a = target.areas[0]
      const spill = ctx.spillArea(a.sheet, a.r0, a.c0)
      return spill ? { kind: "ref", areas: [spill] } : REF
    }
    case "binary":
      return binary(node.op, node.left, node.right, ctx)
  }
}

// ── References ─────────────────────────────────────────────────────────────────────────────────────────────────

export function refArea(node: RefNode, ctx: EvalContext): Area | CellError {
  if (node.invalid) return REF
  let sheet = ctx.sheet
  if (node.sheet !== null) {
    sheet = ctx.sheetIndex(node.sheet)
    if (sheet < 0) return REF
  }
  const { from } = node
  const to = node.to ?? from
  const r0 = from.row ?? 0
  const r1 = to.row ?? MAX_ROWS - 1
  const c0 = from.col ?? 0
  const c1 = to.col ?? MAX_COLS - 1
  return { sheet, r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) }
}

function resolveRef(node: RefNode, ctx: EvalContext): Value {
  if (node.sheetTo !== null) {
    // A 3-D reference: one area per sheet in the span.
    const first = ctx.sheetIndex(node.sheet!)
    const last = ctx.sheetIndex(node.sheetTo)
    if (first < 0 || last < 0) return REF
    const areas: Area[] = []
    for (let s = Math.min(first, last); s <= Math.max(first, last); s++) {
      const area = refArea({ ...node, sheet: ctx.sheetName(s), sheetTo: null }, ctx)
      if (isError(area)) return area
      areas.push(area)
    }
    return { kind: "ref", areas, threeD: true }
  }
  const area = refArea(node, ctx)
  return isError(area) ? area : { kind: "ref", areas: [area] }
}

const structuredSpecs = new Map<string, StructuredSpec | null>()

/** A structured reference's area in the workbook's tables, or the error it gives. */
function resolveStructured(text: string, ctx: EvalContext): Value {
  let spec = structuredSpecs.get(text)
  if (spec === undefined) {
    spec = parseStructured(text)
    if (structuredSpecs.size > 5000) structuredSpecs.clear()
    structuredSpecs.set(text, spec)
  }
  if (!spec) return REF
  const area = structuredArea(spec, ctx.tables?.() ?? [], ctx.sheet, ctx.row, ctx.col)
  return typeof area === "string" ? errorValue(area) : { kind: "ref", areas: [area] }
}

function resolveName(name: string, sheet: string | null, ctx: EvalContext): Value {
  if (sheet === null && ctx.bound) {
    const bound = ctx.bound.get(name.toUpperCase())
    if (bound !== undefined) return bound
  }
  const index = sheet === null ? ctx.sheet : ctx.sheetIndex(sheet)
  if (index < 0) return REF
  const formula = ctx.defined(name, index)
  if (!formula) return NAME
  return evaluate(formula, ctx)
}

/** The bounding area of two references (the : operator). Both must be single areas on the same sheet. */
function rangeOp(a: Value, b: Value): Value {
  if (isError(a)) return a
  if (isError(b)) return b
  if (!isRef(a) || !isRef(b)) return VALUE
  if (a.areas.length !== 1 || b.areas.length !== 1) return VALUE
  const x = a.areas[0]
  const y = b.areas[0]
  if (x.sheet !== y.sheet) return VALUE
  return {
    kind: "ref",
    areas: [
      {
        sheet: x.sheet,
        r0: Math.min(x.r0, y.r0),
        c0: Math.min(x.c0, y.c0),
        r1: Math.max(x.r1, y.r1),
        c1: Math.max(x.c1, y.c1),
      },
    ],
  }
}

function intersectOp(a: Value, b: Value): Value {
  if (isError(a)) return a
  if (isError(b)) return b
  if (!isRef(a) || !isRef(b)) return VALUE
  const areas: Area[] = []
  for (const x of a.areas) {
    for (const y of b.areas) {
      if (x.sheet !== y.sheet) continue
      const r0 = Math.max(x.r0, y.r0)
      const r1 = Math.min(x.r1, y.r1)
      const c0 = Math.max(x.c0, y.c0)
      const c1 = Math.min(x.c1, y.c1)
      if (r0 <= r1 && c0 <= c1) areas.push({ sheet: x.sheet, r0, r1, c0, c1 })
    }
  }
  return areas.length ? { kind: "ref", areas } : NULL
}

function unionOp(a: Value, b: Value): Value {
  if (isError(a)) return a
  if (isError(b)) return b
  if (!isRef(a) || !isRef(b)) return VALUE
  return { kind: "ref", areas: [...a.areas, ...b.areas] }
}

/** @: the one cell of a range in the formula's row or column (legacy implicit intersection). */
export function implicitIntersection(value: Value, ctx: EvalContext): Value {
  if (isMatrix(value)) return value.data[0] ?? VALUE
  if (!isRef(value)) return value
  if (value.areas.length !== 1) return VALUE
  const a = value.areas[0]
  if (a.r0 === a.r1 && a.c0 === a.c1) return ctx.cell(a.sheet, a.r0, a.c0)
  if (a.c0 === a.c1 && ctx.row >= a.r0 && ctx.row <= a.r1 && a.sheet === ctx.sheet)
    return ctx.cell(a.sheet, ctx.row, a.c0)
  if (a.r0 === a.r1 && ctx.col >= a.c0 && ctx.col <= a.c1 && a.sheet === ctx.sheet)
    return ctx.cell(a.sheet, a.r0, ctx.col)
  if (a.r0 === a.r1 && a.c0 === a.c1) return ctx.cell(a.sheet, a.r0, a.c0)
  return VALUE
}

/** Every cell of a (single-area) reference as an array. */
export function refToMatrix(ref: RefValue, ctx: EvalContext): Matrix {
  const a = ref.areas[0]
  const rows = areaRows(a)
  const cols = areaCols(a)
  const data: Scalar[] = new Array(rows * cols).fill(null)
  const used = ctx.usedArea(a)
  if (used) {
    ctx.values(used, (value, row, col) => {
      data[(row - a.r0) * cols + (col - a.c0)] = value
    })
  }
  return matrix(rows, cols, data)
}

function mapMatrix(m: Matrix, f: (v: Scalar) => Scalar): Matrix {
  return matrix(m.rows, m.cols, m.data.map(f))
}

// ── Operators ──────────────────────────────────────────────────────────────────────────────────────────────────

/** A value as an array for element-wise work: references read their cells, scalars are 1×1. */
export function toMatrix(value: Value, ctx: EvalContext): Matrix | CellError {
  if (isMatrix(value)) return value
  if (isRef(value)) {
    if (value.areas.length !== 1) return VALUE
    return refToMatrix(value, ctx)
  }
  return matrix(1, 1, [value])
}

/** A value used where one scalar is wanted; multi-cell values are lifted by the caller instead. */
export function toScalar(value: Value, ctx: EvalContext): Scalar {
  if (isRef(value)) {
    if (value.areas.length !== 1) return VALUE
    const a = value.areas[0]
    if (a.r0 === a.r1 && a.c0 === a.c1) return ctx.cell(a.sheet, a.r0, a.c0)
    return implicitIntersection(value, ctx) as Scalar
  }
  if (isMatrix(value)) return value.data[0] ?? VALUE
  return value
}

export function isSingle(value: Value): boolean {
  if (isRef(value)) {
    const a = value.areas[0]
    return value.areas.length === 1 && a.r0 === a.r1 && a.c0 === a.c1
  }
  if (isMatrix(value)) return value.rows === 1 && value.cols === 1
  return true
}

/** Applies a scalar function to every element (a reference or array argument spills). */
export function lift1(value: Value, ctx: EvalContext, f: (v: Scalar) => Scalar): Value {
  if (isSingle(value)) return f(toScalar(value, ctx))
  const m = toMatrix(value, ctx)
  if (isError(m)) return m
  return mapMatrix(m, f)
}

/**
 * Element-wise over two values with Excel's broadcasting: a single value or a single row/column stretches to the
 * other's size; where sizes otherwise differ the extra positions are #N/A.
 */
export function lift2(a: Value, b: Value, ctx: EvalContext, f: (x: Scalar, y: Scalar) => Scalar): Value {
  if (isSingle(a) && isSingle(b)) return f(toScalar(a, ctx), toScalar(b, ctx))
  const ma = toMatrix(a, ctx)
  if (isError(ma)) return ma
  const mb = toMatrix(b, ctx)
  if (isError(mb)) return mb
  return broadcast([ma, mb], (xs) => f(xs[0], xs[1]))
}

export function broadcast(ms: Matrix[], f: (xs: Scalar[]) => Scalar): Matrix {
  const rows = Math.max(...ms.map((m) => m.rows))
  const cols = Math.max(...ms.map((m) => m.cols))
  const data: Scalar[] = new Array(rows * cols)
  const xs: Scalar[] = new Array(ms.length)
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let missing = false
      for (let k = 0; k < ms.length; k++) {
        const m = ms[k]
        const rr = m.rows === 1 ? 0 : r
        const cc = m.cols === 1 ? 0 : c
        if (rr >= m.rows || cc >= m.cols) {
          missing = true
          break
        }
        xs[k] = m.data[rr * m.cols + cc]
      }
      data[r * cols + c] = missing ? NA : f(xs)
    }
  }
  return matrix(rows, cols, data)
}

/**
 * Excel's cancellation clean-up: when a sum or difference is tiny compared with its operands (the last bits of
 * binary floating point, below 2^-49 of the larger), the result is 0. =0.1+0.2-0.3 shows 0, not 5.55E-17. Excel
 * does it only for the formula's top-level + or − and for SUM's last addition (checked in Excel).
 */
export function cleanSum(result: number, a: number, b: number): number {
  if (result === 0) return 0
  const scale = Math.max(Math.abs(a), Math.abs(b))
  return Math.abs(result) < scale * 2 ** -49 ? 0 : result
}

function arithmetic(op: BinaryOp, x: Scalar, y: Scalar, snap = false): Scalar {
  const a = toNumber(x)
  if (isError(a)) return a
  const b = toNumber(y)
  if (isError(b)) return b
  let r: number
  switch (op) {
    case "+":
      r = snap ? cleanSum(a + b, a, b) : a + b
      break
    case "-":
      r = snap ? cleanSum(a - b, a, b) : a - b
      break
    case "*":
      r = a * b
      break
    case "/":
      if (b === 0) return DIV0
      r = a / b
      break
    case "^":
      return power(a, b)
    default:
      return VALUE
  }
  return Number.isFinite(r) ? r : NUM
}

/**
 * x^n for a whole n by right-to-left repeated squaring, which is how Excel computes it: 1.07^30 and
 * (1+0.05/12)^360 then agree with Excel to the last bit, where Math.pow differs in the 15th digit.
 */
export function squarePower(x: number, n: number): number {
  if (n < 0) return 1 / squarePower(x, -n)
  let r = 1
  for (let b = x; n > 0; n = Math.floor(n / 2), b *= b) if (n % 2 === 1) r *= b
  return r
}

/**
 * x^y as Excel computes it: 0^0 is #NUM!, 0^-n is #DIV/0!, and a negative base takes an odd root ((-8)^(1/3) is
 * -2, computed as -e^(ln 8 / 3) like Excel, which gives -1.9999999999999998).
 */
export function power(x: number, y: number): number | CellError {
  if (x === 0 && y === 0) return NUM
  if (x === 0 && y < 0) return DIV0
  let r = Number.isInteger(y) && Math.abs(y) < 2 ** 31 ? squarePower(x, y) : x ** y
  if (x < 0 && !Number.isInteger(y)) {
    const n = Math.round(1 / y)
    r = Math.abs(1 / y - n) < 1e-12 && n % 2 !== 0 ? -Math.exp(Math.log(-x) * y) : NaN
  }
  return Number.isFinite(r) ? r : NUM
}

function comparison(op: BinaryOp, x: Scalar, y: Scalar): Scalar {
  if (isError(x)) return x
  if (isError(y)) return y
  const c = compareScalars(x, y)
  switch (op) {
    case "=":
      return c === 0
    case "<>":
      return c !== 0
    case "<":
      return c < 0
    case ">":
      return c > 0
    case "<=":
      return c <= 0
    default:
      return c >= 0
  }
}

function concat(x: Scalar, y: Scalar): Scalar {
  const a = toText(x)
  if (isError(a)) return a
  const b = toText(y)
  if (isError(b)) return b
  return a + b
}

function binary(op: BinaryOp, leftNode: Node, rightNode: Node, ctx: EvalContext, snap = false): Value {
  const left = evaluate(leftNode, ctx)
  const right = evaluate(rightNode, ctx)
  switch (op) {
    case ":":
      return rangeOp(left, right)
    case " ":
      return intersectOp(left, right)
    case ",":
      return unionOp(left, right)
    case "&":
      return lift2(left, right, ctx, concat)
    case "=":
    case "<>":
    case "<":
    case ">":
    case "<=":
    case ">=":
      return lift2(left, right, ctx, (x, y) => comparison(op, x, y))
    default:
      return lift2(left, right, ctx, (x, y) => arithmetic(op, x, y, snap))
  }
}

// ── Functions ──────────────────────────────────────────────────────────────────────────────────────────────────

/** A lazily evaluated argument, for IF, IFERROR, CHOOSE and friends that skip branches. */
export type Thunk = () => Value

function callFunction(name: string, args: Node[], ctx: EvalContext): Value {
  const def = FUNCTIONS.get(name)
  if (!def) return NAME
  if (args.length < def.min || args.length > def.max) return VALUE
  if (def.omittedIsNA && args.some((arg) => arg.type === "missing")) return NA
  if (def.lazy)
    return def.lazy(
      args.map((arg) => () => evaluate(arg, ctx)),
      ctx,
      args,
    )
  const values = args.map((arg) => evaluate(arg, ctx))
  return callEager(def, values, ctx)
}

export function callEager(def: FnDef, values: Value[], ctx: EvalContext): Value {
  const kinds = def.params ?? []
  const kindAt = (i: number) => kinds[Math.min(i, kinds.length - 1)] ?? "any"
  // Scalar parameters given an array or a multi-cell range lift the whole call over the elements.
  const lifted: number[] = []
  for (let i = 0; i < values.length; i++) {
    if (kindAt(i) === "scalar" && !isSingle(values[i])) lifted.push(i)
  }
  if (lifted.length === 0) {
    const args = values.map((v, i) => (kindAt(i) === "scalar" ? toScalar(v, ctx) : v))
    if (!def.errors) {
      for (let i = 0; i < args.length; i++) {
        if (kindAt(i) === "scalar" && isError(args[i])) return args[i]
      }
    }
    return def.call!(args, ctx)
  }
  const ms: Matrix[] = []
  for (const i of lifted) {
    const m = toMatrix(values[i], ctx)
    if (isError(m)) return m
    ms.push(m)
  }
  return broadcast(ms, (xs) => {
    const args = values.map((v, i) => {
      const k = lifted.indexOf(i)
      return k >= 0 ? xs[k] : kindAt(i) === "scalar" ? toScalar(v, ctx) : v
    })
    if (!def.errors) {
      for (let i = 0; i < args.length; i++) {
        if (kindAt(i) === "scalar" && isError(args[i])) return args[i] as Scalar
      }
    }
    const result = def.call!(args, ctx)
    return isRef(result) || isMatrix(result) ? toScalar(result, ctx) : result
  })
}

export function errorOf(code: ErrorCode): CellError {
  return errorValue(code)
}

export { DIV0, NA, NAME, NULL, NUM, REF, VALUE }
