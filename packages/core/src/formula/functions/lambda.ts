/**
 * LAMBDA and the functions that take one: MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, and ISOMITTED for a parameter
 * left out of a call. A LAMBDA keeps the LET names around it (by reference, so a LET name can call itself), and the
 * helpers call it through `invoke`, which also takes a built-in function named without brackets.
 */
import { isError, type CellError } from "../../values"
import type { Node } from "../ast"
import { CALC, NA, VALUE } from "../coerce"
import type { EvalContext } from "../context"
import { invoke, OMITTED_KEY, toMatrix, toScalar } from "../evaluate"
import { isLambda, isMatrix, isRef, matrix, type LambdaValue, type Matrix, type Scalar, type Value } from "../types"
import { define } from "./registry"

define("LAMBDA", {
  min: 1,
  max: 254,
  lazy: (_args, ctx, nodes) => {
    const params: string[] = []
    for (const n of nodes.slice(0, -1)) {
      if (n.type !== "name" || n.sheet !== null) return VALUE
      const name = n.name.toUpperCase().replace(/^_XLPM\./, "")
      if (params.includes(name)) return VALUE
      params.push(name)
    }
    const body = nodes[nodes.length - 1]
    if (body.type === "missing") return VALUE
    return { kind: "lambda", params, body, bound: ctx.bound } satisfies LambdaValue
  },
})

define("ISOMITTED", {
  min: 1,
  max: 1,
  lazy: (_args, ctx, nodes) => {
    const n: Node = nodes[0]
    if (n.type !== "name" || n.sheet !== null) return false
    return ctx.bound?.get(OMITTED_KEY + n.name.toUpperCase()) === true
  },
})

/** The function argument of a helper: a LAMBDA (or built-in) taking `count` parameters, or the error to give. */
function fnArg(v: Value | undefined, count: number): LambdaValue | CellError {
  if (v === undefined || v === null) return CALC
  if (isError(v)) return v
  if (!isLambda(v)) return VALUE
  if (!v.builtin && v.params.length !== count) return VALUE
  return v
}

/** A helper's result for one element: one value (an array or range there is #CALC!). */
function single(v: Value, ctx: EvalContext): Scalar | CellError {
  if (isLambda(v)) return CALC
  if (isMatrix(v)) return v.rows === 1 && v.cols === 1 ? v.data[0] : CALC
  if (isRef(v)) {
    const a = v.areas[0]
    return v.areas.length === 1 && a.r0 === a.r1 && a.c0 === a.c1 ? toScalar(v, ctx) : CALC
  }
  return v
}

/** An array argument's values (a range reads its cells; a single value is 1×1). */
function arrayOf(v: Value, ctx: EvalContext): Matrix | CellError {
  return toMatrix(v, ctx)
}

/** One element as a lambda argument: a range's cell stays a reference, so ISBLANK and the like see the cell. */
function elementArg(source: Value, m: Matrix, r: number, c: number): Value {
  if (isRef(source) && source.areas.length === 1) {
    const a = source.areas[0]
    const area = { sheet: a.sheet, r0: a.r0 + r, r1: a.r0 + r, c0: a.c0 + c, c1: a.c0 + c }
    return { kind: "ref", areas: [area] }
  }
  return m.data[r * m.cols + c]
}

define("MAP", {
  min: 2,
  max: 254,
  call: (args, ctx) => {
    const fn = fnArg(args[args.length - 1], args.length - 1)
    if (isError(fn)) return fn
    const sources = args.slice(0, -1)
    const arrays: Matrix[] = []
    for (const s of sources) {
      const m = arrayOf(s, ctx)
      if (isError(m)) return m
      arrays.push(m)
    }
    // Arrays of different sizes give #N/A where any of them has no element.
    const rows = Math.max(...arrays.map((m) => m.rows))
    const cols = Math.max(...arrays.map((m) => m.cols))
    const data: Scalar[] = []
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (arrays.some((m) => r >= m.rows || c >= m.cols)) {
          data.push(NA)
          continue
        }
        const v = single(
          invoke(
            fn,
            arrays.map((m, i) => elementArg(sources[i], m, r, c)),
            ctx,
          ),
          ctx,
        )
        if (v === CALC) return CALC
        data.push(v)
      }
    }
    return matrix(rows, cols, data)
  },
})

/** REDUCE and SCAN: the accumulator through every element, row by row. */
function fold(args: Value[], ctx: EvalContext, keep: boolean): Value {
  const fn = fnArg(args[2], 2)
  if (isError(fn)) return fn
  const m = arrayOf(args[1], ctx)
  if (isError(m)) return m
  let acc: Value = args[0] === undefined ? null : args[0]
  const steps: Scalar[] = []
  for (let r = 0; r < m.rows; r++) {
    for (let c = 0; c < m.cols; c++) {
      acc = invoke(fn, [acc, elementArg(args[1], m, r, c)], ctx)
      if (keep) {
        const v = single(acc, ctx)
        if (v === CALC) return CALC
        steps.push(v)
      }
    }
  }
  if (keep) return matrix(m.rows, m.cols, steps)
  return isLambda(acc) ? CALC : acc
}

define("REDUCE", { min: 2, max: 3, call: (args, ctx) => fold(args, ctx, false) })
define("SCAN", { min: 2, max: 3, call: (args, ctx) => fold(args, ctx, true) })

for (const byRow of [true, false]) {
  define(byRow ? "BYROW" : "BYCOL", {
    min: 1,
    max: 2,
    call: (args, ctx) => {
      const fn = fnArg(args[1], 1)
      if (isError(fn)) return fn
      const m = arrayOf(args[0], ctx)
      if (isError(m)) return m
      const count = byRow ? m.rows : m.cols
      const out: Scalar[] = []
      for (let i = 0; i < count; i++) {
        const part = byRow
          ? matrix(1, m.cols, m.data.slice(i * m.cols, (i + 1) * m.cols))
          : matrix(
              m.rows,
              1,
              Array.from({ length: m.rows }, (_, r) => m.data[r * m.cols + i]),
            )
        const v = single(invoke(fn, [part], ctx), ctx)
        if (v === CALC) return CALC
        out.push(v)
      }
      return byRow ? matrix(count, 1, out) : matrix(1, count, out)
    },
  })
}

define("MAKEARRAY", {
  min: 3,
  max: 3,
  call: (args, ctx) => {
    const size = (v: Value): number | CellError => {
      const s = toScalar(v, ctx)
      if (isError(s)) return s
      return typeof s === "number" && Number.isInteger(s) && s >= 1 ? s : VALUE
    }
    const rows = size(args[0])
    if (isError(rows)) return rows
    const cols = size(args[1])
    if (isError(cols)) return cols
    const fn = fnArg(args[2], 2)
    if (isError(fn)) return fn
    if (rows * cols > 1_048_576 * 16_384) return VALUE
    const data: Scalar[] = []
    for (let r = 1; r <= rows; r++) {
      for (let c = 1; c <= cols; c++) {
        const v = single(invoke(fn, [r, c], ctx), ctx)
        if (v === CALC) return CALC
        data.push(v)
      }
    }
    return matrix(rows, cols, data)
  },
})
