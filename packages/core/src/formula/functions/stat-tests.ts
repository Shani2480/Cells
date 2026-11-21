/**
 * Statistical tests (Z.TEST, T.TEST, F.TEST, CHISQ.TEST and their old names) and the descriptive statistics that
 * look at a whole data set: KURT, SKEW, TRIMMEAN, PERCENTRANK, PROB, MODE.MULT and FREQUENCY.
 */
import { isError, type CellError } from "../../values"
import { DIV0, NA, NUM, VALUE } from "../coerce"
import { matrix, type Scalar, type Value } from "../types"
import { betaI, gammaQ, normCdf } from "./special"
import { arrayArg, collectNumbers, int, num } from "./helpers"
import { define } from "./registry"

type Ctx = Parameters<typeof collectNumbers>[1]

function mean(xs: number[]) {
  let s = 0
  for (const x of xs) s += x
  return s / xs.length
}
/** Σ(x − mean)² */
function squares(xs: number[], m = mean(xs)) {
  let s = 0
  for (const x of xs) s += (x - m) ** 2
  return s
}

// ── Tests ───────────────────────────────────────────────────────────────────────────────────────────────────────

function zTest(a: Value[], ctx: Ctx): Scalar {
  const xs = collectNumbers([a[0]], ctx)
  if (isError(xs)) return xs
  const x = num(a[1])
  if (isError(x)) return x
  if (!xs.length) return NA
  let sigma: number
  if (a[2] != null) {
    const s = num(a[2])
    if (isError(s)) return s
    sigma = s
  } else {
    if (xs.length < 2) return DIV0
    sigma = Math.sqrt(squares(xs) / (xs.length - 1))
  }
  if (sigma === 0) return DIV0
  return 1 - normCdf((mean(xs) - x) / (sigma / Math.sqrt(xs.length)))
}
for (const name of ["Z.TEST", "ZTEST"])
  define(name, { min: 2, max: 3, params: ["any", "scalar", "scalar"], call: zTest })

/** The two-tailed or one-tailed probability of Student's t with df degrees of freedom. */
function tProbability(t: number, df: number, tails: number): number {
  const upper = 0.5 * betaI(df / (df + t * t), df / 2, 0.5)
  return tails * upper
}

/** Numbers of an array at each position (null where it holds no number), for T.TEST's pairs. */
function positions(v: Value, ctx: Ctx): (number | null)[] | CellError {
  const m = arrayArg(v, ctx)
  if (isError(m)) return m
  const out: (number | null)[] = []
  for (const x of m.data) {
    if (isError(x)) return x
    out.push(typeof x === "number" ? x : null)
  }
  return out
}

function tTest(a: Value[], ctx: Ctx): Scalar {
  const tailsN = int(a[2])
  const typeN = int(a[3])
  if (isError(tailsN)) return tailsN
  if (isError(typeN)) return typeN
  if ((tailsN !== 1 && tailsN !== 2) || typeN < 1 || typeN > 3) return NUM
  if (typeN === 1) {
    const p = positions(a[0], ctx)
    const q = positions(a[1], ctx)
    if (isError(p)) return p
    if (isError(q)) return q
    if (p.length !== q.length) return NA
    const d: number[] = []
    for (let i = 0; i < p.length; i++) {
      const x = p[i]
      const y = q[i]
      if (x !== null && y !== null) d.push(x - y)
    }
    if (d.length < 2) return DIV0
    const sd = Math.sqrt(squares(d) / (d.length - 1))
    if (sd === 0) return DIV0
    const t = Math.abs(mean(d)) / (sd / Math.sqrt(d.length))
    return tProbability(t, d.length - 1, tailsN)
  }
  const x = collectNumbers([a[0]], ctx)
  const y = collectNumbers([a[1]], ctx)
  if (isError(x)) return x
  if (isError(y)) return y
  const [n1, n2] = [x.length, y.length]
  if (n1 < 2 || n2 < 2) return DIV0
  const [v1, v2] = [squares(x) / (n1 - 1), squares(y) / (n2 - 1)]
  const diff = Math.abs(mean(x) - mean(y))
  if (typeN === 2) {
    const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2)
    const se = Math.sqrt(pooled * (1 / n1 + 1 / n2))
    if (se === 0) return DIV0
    return tProbability(diff / se, n1 + n2 - 2, tailsN)
  }
  const [s1, s2] = [v1 / n1, v2 / n2]
  const se = Math.sqrt(s1 + s2)
  if (se === 0) return DIV0
  // Welch–Satterthwaite degrees of freedom (not rounded, as in Excel 2010 and later).
  const df = (s1 + s2) ** 2 / (s1 ** 2 / (n1 - 1) + s2 ** 2 / (n2 - 1))
  return tProbability(diff / se, df, tailsN)
}
for (const name of ["T.TEST", "TTEST"])
  define(name, { min: 4, max: 4, params: ["any", "any", "scalar", "scalar"], call: tTest })

function fTest(a: Value[], ctx: Ctx): Scalar {
  const x = collectNumbers([a[0]], ctx)
  const y = collectNumbers([a[1]], ctx)
  if (isError(x)) return x
  if (isError(y)) return y
  if (x.length < 2 || y.length < 2) return DIV0
  const v1 = squares(x) / (x.length - 1)
  const v2 = squares(y) / (y.length - 1)
  if (v1 === 0 || v2 === 0) return DIV0
  const [d1, d2] = [x.length - 1, y.length - 1]
  const f = v1 / v2
  // Two-tailed: twice the smaller tail of F(d1, d2).
  const lower = betaI((d1 * f) / (d1 * f + d2), d1 / 2, d2 / 2)
  return Math.min(1, 2 * Math.min(lower, 1 - lower))
}
for (const name of ["F.TEST", "FTEST"]) define(name, { min: 2, max: 2, call: fTest })

function chisqTest(a: Value[], ctx: Ctx): Scalar {
  const actual = arrayArg(a[0], ctx)
  const expected = arrayArg(a[1], ctx)
  if (isError(actual)) return actual
  if (isError(expected)) return expected
  if (actual.rows !== expected.rows || actual.cols !== expected.cols) return NA
  if (actual.rows * actual.cols < 2) return NA
  let chi = 0
  for (let i = 0; i < actual.data.length; i++) {
    const o = actual.data[i]
    const e = expected.data[i]
    if (isError(o)) return o
    if (isError(e)) return e
    if (typeof o !== "number" || typeof e !== "number") continue
    if (e === 0) return DIV0
    if (e < 0) return NUM
    chi += (o - e) ** 2 / e
  }
  const { rows, cols } = actual
  const df = rows > 1 && cols > 1 ? (rows - 1) * (cols - 1) : Math.max(rows, cols) - 1
  return gammaQ(df / 2, chi / 2)
}
for (const name of ["CHISQ.TEST", "CHITEST"]) define(name, { min: 2, max: 2, call: chisqTest })

// ── Shape of a data set ─────────────────────────────────────────────────────────────────────────────────────────

function numbersOf(a: Value[], ctx: Ctx, f: (xs: number[]) => Scalar) {
  const xs = collectNumbers(a, ctx)
  return isError(xs) ? xs : f(xs)
}
define("KURT", {
  min: 1,
  max: 255,
  call: (a, ctx) =>
    numbersOf(a, ctx, (xs) => {
      const n = xs.length
      if (n < 4) return DIV0
      const m = mean(xs)
      const s = Math.sqrt(squares(xs, m) / (n - 1))
      if (s === 0) return DIV0
      let sum = 0
      for (const x of xs) sum += ((x - m) / s) ** 4
      return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * sum - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
    }),
})
define("SKEW", {
  min: 1,
  max: 255,
  call: (a, ctx) =>
    numbersOf(a, ctx, (xs) => {
      const n = xs.length
      if (n < 3) return DIV0
      const m = mean(xs)
      const s = Math.sqrt(squares(xs, m) / (n - 1))
      if (s === 0) return DIV0
      let sum = 0
      for (const x of xs) sum += ((x - m) / s) ** 3
      return (n / ((n - 1) * (n - 2))) * sum
    }),
})
define("SKEW.P", {
  min: 1,
  max: 255,
  call: (a, ctx) =>
    numbersOf(a, ctx, (xs) => {
      const n = xs.length
      if (n < 1) return DIV0
      const m = mean(xs)
      const s = Math.sqrt(squares(xs, m) / n)
      if (s === 0) return DIV0
      let sum = 0
      for (const x of xs) sum += ((x - m) / s) ** 3
      return sum / n
    }),
})

define("TRIMMEAN", {
  min: 2,
  max: 2,
  params: ["any", "scalar"],
  call: (a, ctx) => {
    const xs = collectNumbers([a[0]], ctx)
    if (isError(xs)) return xs
    const p = num(a[1])
    if (isError(p)) return p
    if (p < 0 || p >= 1 || !xs.length) return NUM
    // Points left out, rounded down to an even number, half from each end.
    const k = Math.floor((xs.length * p) / 2)
    const sorted = [...xs].sort((x, y) => x - y).slice(k, xs.length - k)
    return mean(sorted)
  },
})

/** PERCENTRANK: where x falls in the sorted data, interpolated, cut (not rounded) to `digits` digits. */
function percentRank(exclusive: boolean) {
  return (a: Value[], ctx: Ctx): Scalar => {
    const xs = collectNumbers([a[0]], ctx)
    if (isError(xs)) return xs
    const x = num(a[1])
    if (isError(x)) return x
    const digits = a[2] == null ? 3 : int(a[2])
    if (isError(digits)) return digits
    if (digits < 1) return NUM
    const n = xs.length
    if (!n) return NUM
    const s = [...xs].sort((p, q) => p - q)
    if (x < s[0] || x > s[n - 1]) return NA
    let below = 0
    while (below < n && s[below] < x) below++
    let position: number
    if (below < n && s[below] === x) position = below
    else position = below - 1 + (x - s[below - 1]) / (s[below] - s[below - 1])
    let rank: number
    if (exclusive) rank = (position + 1) / (n + 1)
    else if (n === 1) rank = 1
    else rank = position / (n - 1)
    const scale = 10 ** digits
    // Cut to the digits asked for; the small nudge keeps 0.7 from becoming 0.699 through binary rounding.
    return Math.floor(rank * scale * (1 + 1e-13)) / scale
  }
}
for (const [name, exclusive] of [
  ["PERCENTRANK", false],
  ["PERCENTRANK.INC", false],
  ["PERCENTRANK.EXC", true],
] as const) {
  define(name, { min: 2, max: 3, params: ["any", "scalar", "scalar"], call: percentRank(exclusive) })
}

define("PROB", {
  min: 3,
  max: 4,
  params: ["any", "any", "scalar", "scalar"],
  call: (a, ctx) => {
    const xs = arrayArg(a[0], ctx)
    const ps = arrayArg(a[1], ctx)
    if (isError(xs)) return xs
    if (isError(ps)) return ps
    if (xs.data.length !== ps.data.length) return NA
    const lower = num(a[2])
    if (isError(lower)) return lower
    const upper = a[3] == null ? lower : num(a[3])
    if (isError(upper)) return upper
    let total = 0
    let inside = 0
    for (let i = 0; i < xs.data.length; i++) {
      const x = xs.data[i]
      const p = ps.data[i]
      if (isError(x)) return x
      if (isError(p)) return p
      if (typeof p !== "number") continue
      if (p <= 0 || p > 1) return NUM
      total += p
      if (typeof x === "number" && x >= lower && x <= upper) inside += p
    }
    if (Math.abs(total - 1) > 1e-6) return NUM
    return inside
  },
})

define("MODE.MULT", {
  min: 1,
  max: 255,
  call: (a, ctx) => {
    const xs = collectNumbers(a, ctx)
    if (isError(xs)) return xs
    const counts = new Map<number, number>()
    for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1)
    let best = 1
    for (const c of counts.values()) best = Math.max(best, c)
    if (best < 2) return NA
    const modes: number[] = []
    for (const [x, c] of counts) if (c === best) modes.push(x)
    return matrix(modes.length, 1, modes)
  },
})

define("FREQUENCY", {
  min: 2,
  max: 2,
  call: (a, ctx) => {
    const data = collectNumbers([a[0]], ctx)
    const bins = collectNumbers([a[1]], ctx)
    if (isError(data)) return data
    if (isError(bins)) return bins
    // Each bin counts the values above the next smaller bin, up to and including itself; one more for the rest.
    const order = bins.map((b, i) => [b, i] as const).sort((p, q) => p[0] - q[0])
    const counts = new Array<number>(bins.length + 1).fill(0)
    for (const x of data) {
      const at = order.findIndex(([b]) => x <= b)
      counts[at < 0 ? bins.length : order[at][1]]++
    }
    return matrix(counts.length, 1, counts)
  },
})

export { VALUE }
