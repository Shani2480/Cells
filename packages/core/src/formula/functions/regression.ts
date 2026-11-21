/**
 * Least-squares fits: LINEST and LOGEST (the coefficients, and with stats=TRUE the table of standard errors, r², F
 * and sums of squares), TREND and GROWTH (the fitted values for new x). Solved with Householder QR, dropping x columns
 * that are a combination of others (their coefficient shows 0, as in Excel).
 */
import { isError, type CellError } from "../../values"
import { NA, NUM, REF, toBool, VALUE } from "../coerce"
import { matrix, type Matrix, type Scalar, type Value } from "../types"
import { arrayArg } from "./helpers"
import { define } from "./registry"

type Ctx = Parameters<typeof arrayArg>[1]

/** A numeric matrix (any text, logical or blank is #VALUE!). */
function numericMatrix(v: Value, ctx: Ctx): { rows: number; cols: number; data: number[] } | CellError {
  const m = arrayArg(v, ctx)
  if (isError(m)) return m
  const data: number[] = []
  for (const x of m.data) {
    if (isError(x)) return x
    if (typeof x !== "number") return VALUE
    data.push(x)
  }
  return { rows: m.rows, cols: m.cols, data }
}

type Design = {
  y: number[]
  /** One array per x variable, n values each. */
  xs: number[][]
  /** y is a row (variables then run down the rows of known_x). */
  byRow: boolean
}

/** known_y and known_x as a vector of n values and k variables of n values (Excel's shape rules). */
function design(yArg: Value, xArg: Value | undefined, ctx: Ctx): Design | CellError {
  const y = numericMatrix(yArg, ctx)
  if (isError(y)) return y
  const n = y.data.length
  if (xArg == null) {
    return { y: y.data, xs: [Array.from({ length: n }, (_, i) => i + 1)], byRow: y.rows === 1 && y.cols > 1 }
  }
  const x = numericMatrix(xArg, ctx)
  if (isError(x)) return x
  if (y.rows === 1 || y.cols === 1) {
    const byRow = y.rows === 1 && y.cols > 1
    if (byRow ? x.cols === n : x.rows === n) {
      // Each column (or, for a row of y, each row) of known_x is one variable.
      const k = byRow ? x.rows : x.cols
      const xs = Array.from({ length: k }, (_, j) =>
        Array.from({ length: n }, (_, i) => (byRow ? x.data[j * x.cols + i] : x.data[i * x.cols + j])),
      )
      return { y: y.data, xs, byRow }
    }
    if (x.data.length === n) return { y: y.data, xs: [x.data], byRow }
    return REF
  }
  // A y of several rows and columns: one x of the same shape.
  if (x.rows !== y.rows || x.cols !== y.cols) return REF
  return { y: y.data, xs: [x.data], byRow: false }
}

type Fit = {
  /** Coefficients of each x variable (0 for dropped ones). */
  m: number[]
  b: number
  dropped: boolean[]
  /** Standard errors, when the stats are asked for. */
  se?: number[]
  seB?: number
  r2: number
  seY: number
  f: number
  df: number
  ssreg: number
  ssresid: number
}

/**
 * Least squares of y on the columns xs (and a constant when `constant`), by Householder QR. With a constant the data
 * are centred first (the constant then follows from the means), which keeps the digits of the slopes, r² and the sums
 * of squares; x columns that are a combination of others are dropped (coefficient and standard error 0).
 */
function leastSquares(y: number[], xs: number[][], constant: boolean): Fit | CellError {
  const n = y.length
  const k = xs.length
  const meanOf = (v: number[]) => v.reduce((s, x) => s + x, 0) / n
  const yMean = constant ? meanOf(y) : 0
  const xMeans = xs.map((x) => (constant ? meanOf(x) : 0))
  const work = xs.map((x, j) => x.map((v) => v - xMeans[j]))
  const yy = y.map((v) => v - yMean)
  const order: number[] = []
  let row = 0
  for (let j = 0; j < k && row < n; j++) {
    const col = work[j]
    let norm = 0
    for (let i = row; i < n; i++) norm += col[i] * col[i]
    norm = Math.sqrt(norm)
    let scale = 0
    for (let i = 0; i < n; i++) scale = Math.max(scale, Math.abs(xs[j][i]))
    if (norm === 0 || norm <= scale * 1e-13 * Math.sqrt(n)) continue
    const alpha = col[row] > 0 ? -norm : norm
    const v = new Array<number>(n).fill(0)
    for (let i = row; i < n; i++) v[i] = col[i]
    v[row] -= alpha
    let vv = 0
    for (let i = row; i < n; i++) vv += v[i] * v[i]
    const reflect = (c: number[]) => {
      let dot = 0
      for (let i = row; i < n; i++) dot += v[i] * c[i]
      const f = (2 * dot) / vv
      for (let i = row; i < n; i++) c[i] -= f * v[i]
    }
    for (let q = j; q < k; q++) reflect(work[q])
    reflect(yy)
    order.push(j)
    row++
  }
  const r = order.length
  const R = (a: number, b: number) => work[order[b]][a]
  // R c = Qᵀy.
  const coef = new Array<number>(r).fill(0)
  for (let a = r - 1; a >= 0; a--) {
    let s = yy[a]
    for (let b = a + 1; b < r; b++) s -= R(a, b) * coef[b]
    coef[a] = s / R(a, a)
  }
  const m = new Array<number>(k).fill(0)
  order.forEach((j, a) => (m[j] = coef[a]))
  let b = 0
  if (constant) {
    b = yMean
    for (let j = 0; j < k; j++) b -= m[j] * xMeans[j]
  }
  let ssresid = 0
  for (let i = r; i < n; i++) ssresid += yy[i] * yy[i]
  const sstot = constant ? y.reduce((s, v) => s + (v - yMean) ** 2, 0) : y.reduce((s, v) => s + v * v, 0)
  const ssreg = sstot - ssresid
  const df = n - r - (constant ? 1 : 0)
  const r2 = sstot === 0 ? 1 : ssreg / sstot
  const sigma2 = df > 0 ? ssresid / df : 0
  // Standard errors from R⁻¹: var(m) = σ²·diag(R⁻¹R⁻ᵀ); var(b) = σ²·(1/n + ‖R⁻ᵀx̄‖²).
  const rinv: number[][] = Array.from({ length: r }, () => new Array<number>(r).fill(0))
  for (let c = 0; c < r; c++) {
    rinv[c][c] = 1 / R(c, c)
    for (let a = c - 1; a >= 0; a--) {
      let s = 0
      for (let q = a + 1; q <= c; q++) s += R(a, q) * rinv[q][c]
      rinv[a][c] = -s / R(a, a)
    }
  }
  const se = new Array<number>(k).fill(0)
  order.forEach((j, a) => {
    let s = 0
    for (let c = 0; c < r; c++) s += rinv[a][c] ** 2
    se[j] = Math.sqrt(s * sigma2)
  })
  let seB: number | undefined
  if (constant) {
    // w = R⁻ᵀ x̄ (forward substitution with Rᵀ).
    const w = new Array<number>(r).fill(0)
    for (let a = 0; a < r; a++) {
      let s = xMeans[order[a]]
      for (let q = 0; q < a; q++) s -= R(q, a) * w[q]
      w[a] = s / R(a, a)
    }
    seB = Math.sqrt(sigma2 * (1 / n + w.reduce((s, v) => s + v * v, 0)))
  }
  const dropped = xs.map((_, j) => !order.includes(j))
  return {
    m,
    b,
    dropped,
    se,
    seB,
    r2,
    seY: Math.sqrt(sigma2),
    f: df > 0 && r > 0 ? ssreg / r / sigma2 : NaN,
    df,
    ssreg,
    ssresid,
  }
}

function flag(v: Value | undefined, otherwise: boolean): boolean | CellError {
  if (v === null || v === undefined) return otherwise
  if (typeof v === "object" && !isError(v as Scalar)) return VALUE
  return toBool(v as Scalar)
}

function linest(a: Value[], ctx: Ctx, log: boolean): Scalar | Matrix {
  const d = design(a[0], a[1] ?? undefined, ctx)
  if (isError(d)) return d
  const constant = flag(a[2], true)
  const stats = flag(a[3], false)
  if (isError(constant)) return constant
  if (isError(stats)) return stats
  let y = d.y
  if (log) {
    if (y.some((v) => v <= 0)) return NUM
    y = y.map(Math.log)
  }
  const fit = leastSquares(y, d.xs, constant)
  if (isError(fit)) return fit
  const k = d.xs.length
  const cols = k + 1
  const out = (v: number): Scalar => (Number.isFinite(v) ? v : NUM)
  const first: Scalar[] = []
  for (let j = k - 1; j >= 0; j--) first.push(log ? Math.exp(fit.m[j]) : fit.m[j])
  first.push(log ? Math.exp(fit.b) : fit.b)
  if (!stats) return matrix(1, cols, first)
  const rows: Scalar[][] = [first]
  const se: Scalar[] = []
  for (let j = k - 1; j >= 0; j--) se.push(fit.dropped[j] ? 0 : (fit.se?.[j] ?? 0))
  se.push(constant ? (fit.seB ?? 0) : NA)
  rows.push(se)
  const pad = (row: Scalar[]) => [...row, ...new Array<Scalar>(cols - row.length).fill(NA)]
  rows.push(pad([out(fit.r2), out(fit.seY)]))
  rows.push(pad([fit.df === 0 ? NUM : out(fit.f), fit.df]))
  rows.push(pad([out(fit.ssreg), out(fit.ssresid)]))
  return matrix(5, cols, rows.flat())
}

define("LINEST", { min: 1, max: 4, call: (a, ctx) => linest(a, ctx, false) })
define("LOGEST", { min: 1, max: 4, call: (a, ctx) => linest(a, ctx, true) })

/** TREND and GROWTH: the fitted values at new_x, in new_x's shape. */
function trend(a: Value[], ctx: Ctx, log: boolean): Scalar | Matrix {
  const d = design(a[0], a[1] ?? undefined, ctx)
  if (isError(d)) return d
  const constant = flag(a[3], true)
  if (isError(constant)) return constant
  let y = d.y
  if (log) {
    if (y.some((v) => v <= 0)) return NUM
    y = y.map(Math.log)
  }
  const fit = leastSquares(y, d.xs, constant)
  if (isError(fit)) return fit
  const k = d.xs.length
  // new_x: given, or known_x (1…n when that was left out too).
  let rows: number
  let cols: number
  let points: number[][]
  if (a[2] != null) {
    const nx = numericMatrix(a[2], ctx)
    if (isError(nx)) return nx
    if (k === 1) {
      rows = nx.rows
      cols = nx.cols
      points = nx.data.map((v) => [v])
    } else if (d.byRow) {
      if (nx.rows !== k) return REF
      rows = 1
      cols = nx.cols
      points = Array.from({ length: nx.cols }, (_, i) => Array.from({ length: k }, (_, j) => nx.data[j * nx.cols + i]))
    } else {
      if (nx.cols !== k) return REF
      rows = nx.rows
      cols = 1
      points = Array.from({ length: nx.rows }, (_, i) => nx.data.slice(i * k, i * k + k))
    }
  } else {
    const yShape = arrayArg(a[0], ctx)
    if (isError(yShape)) return yShape
    rows = yShape.rows
    cols = yShape.cols
    points = d.y.map((_, i) => d.xs.map((x) => x[i]))
  }
  const values: Scalar[] = points.map((pt) => {
    let s = fit.b
    for (let j = 0; j < k; j++) s += fit.m[j] * pt[j]
    return log ? Math.exp(s) : s
  })
  return rows === 1 && cols === 1 ? values[0] : matrix(rows, cols, values)
}
define("TREND", { min: 1, max: 4, call: (a, ctx) => trend(a, ctx, false) })
define("GROWTH", { min: 1, max: 4, call: (a, ctx) => trend(a, ctx, true) })
