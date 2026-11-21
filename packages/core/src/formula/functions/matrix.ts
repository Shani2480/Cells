/**
 * Matrices (MUNIT, MDETERM, MINVERSE), SERIESSUM, MULTINOMIAL, and the reciprocal and hyperbolic trigonometry Excel
 * 2013 added (ACOT, ACOTH, COTH, CSCH, SECH).
 */
import { isError, type CellError } from "../../values"
import { DIV0, NUM, toNumber, VALUE } from "../coerce"
import type { EvalContext } from "../context"
import { toMatrix } from "../evaluate"
import { matrix, type Matrix, type Value } from "../types"
import { forEachScalar, num } from "./helpers"
import { define, scalarFn } from "./registry"

const n1 = (f: (x: number) => number | CellError) => (a: Value[]) => {
  const x = num(a[0])
  if (isError(x)) return x
  const r = f(x)
  return typeof r === "number" && !Number.isFinite(r) ? NUM : r
}

scalarFn(
  "ACOT",
  1,
  1,
  n1((x) => (x === 0 ? Math.PI / 2 : x > 0 ? Math.atan(1 / x) : Math.PI + Math.atan(1 / x))),
)
scalarFn(
  "ACOTH",
  1,
  1,
  n1((x) => (Math.abs(x) <= 1 ? NUM : Math.atanh(1 / x))),
)
scalarFn(
  "COTH",
  1,
  1,
  n1((x) => (x === 0 ? DIV0 : 1 / Math.tanh(x))),
)
scalarFn(
  "CSCH",
  1,
  1,
  n1((x) => (x === 0 ? DIV0 : 1 / Math.sinh(x))),
)
scalarFn(
  "SECH",
  1,
  1,
  n1((x) => 1 / Math.cosh(x)),
)

/** A numeric matrix argument: every element a number (blanks and text are #VALUE!). */
function numeric(v: Value, ctx: EvalContext): { n: number; m: number; a: number[] } | CellError {
  const mx: Matrix | CellError = toMatrix(v, ctx)
  if (isError(mx)) return mx
  const a: number[] = []
  for (const x of mx.data) {
    if (isError(x)) return x
    if (typeof x !== "number") return VALUE
    a.push(x)
  }
  return { n: mx.rows, m: mx.cols, a }
}

define("MUNIT", {
  min: 1,
  max: 1,
  params: ["scalar"],
  call: (a) => {
    const d = num(a[0])
    if (isError(d)) return d
    const n = Math.trunc(d)
    if (n < 1) return VALUE
    return matrix(
      n,
      n,
      Array.from({ length: n * n }, (_, i) => (Math.floor(i / n) === i % n ? 1 : 0)),
    )
  },
})

/** LU decomposition with partial pivoting: the determinant (0 for a singular matrix). */
function determinant(n: number, src: number[]): number {
  const a = src.slice()
  let det = 1
  for (let k = 0; k < n; k++) {
    let p = k
    for (let i = k + 1; i < n; i++) if (Math.abs(a[i * n + k]) > Math.abs(a[p * n + k])) p = i
    if (a[p * n + k] === 0) return 0
    if (p !== k) {
      for (let j = 0; j < n; j++) [a[k * n + j], a[p * n + j]] = [a[p * n + j], a[k * n + j]]
      det = -det
    }
    const pivot = a[k * n + k]
    det *= pivot
    for (let i = k + 1; i < n; i++) {
      const f = a[i * n + k] / pivot
      if (f === 0) continue
      for (let j = k; j < n; j++) a[i * n + j] -= f * a[k * n + j]
    }
  }
  return det
}

define("MDETERM", {
  min: 1,
  max: 1,
  call: (args, ctx) => {
    const m = numeric(args[0], ctx)
    if (isError(m)) return m
    if (m.n !== m.m) return VALUE
    return determinant(m.n, m.a)
  },
})

define("MINVERSE", {
  min: 1,
  max: 1,
  call: (args, ctx) => {
    const m = numeric(args[0], ctx)
    if (isError(m)) return m
    if (m.n !== m.m) return VALUE
    const n = m.n
    // Gauss–Jordan with partial pivoting on [A | I].
    const a = m.a.slice()
    const inv: number[] = Array.from({ length: n * n }, (_, i) => (Math.floor(i / n) === i % n ? 1 : 0))
    for (let k = 0; k < n; k++) {
      let p = k
      for (let i = k + 1; i < n; i++) if (Math.abs(a[i * n + k]) > Math.abs(a[p * n + k])) p = i
      if (a[p * n + k] === 0) return NUM
      if (p !== k)
        for (let j = 0; j < n; j++) {
          ;[a[k * n + j], a[p * n + j]] = [a[p * n + j], a[k * n + j]]
          ;[inv[k * n + j], inv[p * n + j]] = [inv[p * n + j], inv[k * n + j]]
        }
      const pivot = a[k * n + k]
      for (let j = 0; j < n; j++) {
        a[k * n + j] /= pivot
        inv[k * n + j] /= pivot
      }
      for (let i = 0; i < n; i++) {
        if (i === k) continue
        const f = a[i * n + k]
        if (f === 0) continue
        for (let j = 0; j < n; j++) {
          a[i * n + j] -= f * a[k * n + j]
          inv[i * n + j] -= f * inv[k * n + j]
        }
      }
    }
    return matrix(n, n, inv)
  },
})

define("SERIESSUM", {
  min: 4,
  max: 4,
  call: (args, ctx) => {
    const p: number[] = []
    for (let i = 0; i < 3; i++) {
      const v = num(toScalarOf(args[i], ctx))
      if (isError(v)) return v
      p.push(v)
    }
    const [x, n, step] = p
    const coefficients = toMatrix(args[3], ctx)
    if (isError(coefficients)) return coefficients
    let sum = 0
    let i = 0
    for (const c of coefficients.data) {
      if (isError(c)) return c
      if (c === null) continue
      if (typeof c !== "number") return VALUE
      const e = n + i * step
      if (x === 0 && e <= 0) return NUM
      const power = Math.pow(x, e)
      if (!Number.isFinite(power)) return NUM
      sum += c * power
      i++
    }
    return sum
  },
})

function toScalarOf(v: Value, ctx: EvalContext) {
  const m = toMatrix(v, ctx)
  return isError(m) ? m : (m.data[0] ?? null)
}

define("MULTINOMIAL", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ns: number[] = []
    let error: CellError | null = null
    forEachScalar(args, ctx, (v) => {
      if (isError(v)) error = v
      else if (typeof v === "boolean" || (typeof v === "string" && v === "")) error = VALUE
      else {
        const n = toNumber(v)
        if (isError(n)) error = n
        else ns.push(Math.trunc(n))
      }
      return error === null
    })
    if (error) return error
    if (ns.some((n) => n < 0)) return NUM
    // (Σn)! / Πn!: Excel gives #NUM! once (Σn)! overflows, even when the quotient would fit.
    const total = ns.reduce((s, n) => s + n, 0)
    if (total > 170) return NUM
    let result = 1
    let k = 0
    for (const n of ns) for (let j = 1; j <= n; j++) result = (result * ++k) / j
    return result
  },
})
