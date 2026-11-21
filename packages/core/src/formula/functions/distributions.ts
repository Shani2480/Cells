/**
 * Excel's probability distributions (NORM.DIST, T.DIST, CHISQ.INV, BINOM.DIST…) and their pre-2010 names (NORMDIST,
 * TDIST, CHIINV, BINOMDIST…), with Excel's argument checks: what is truncated to an integer, and what gives #NUM!.
 */
import { isError, type CellError } from "../../values"
import { DIV0, NUM, toBool, VALUE } from "../coerce"
import type { Scalar, Value } from "../types"
import { int, num } from "./helpers"
import { scalarFn } from "./registry"
import {
  betaI,
  betaIc,
  betaPrefix,
  gammaPrefix,
  erfc,
  gamma,
  gammaP,
  gammaQ,
  logBeta,
  logGamma,
  normCdf,
  normInv,
  normPdf,
  solveIncreasing,
} from "./special"

/** Numbers from the arguments, or the first error. */
function numbers(a: Value[], from = 0, to = a.length): number[] | CellError {
  const out: number[] = []
  for (let i = from; i < to; i++) {
    const n = num(a[i])
    if (isError(n)) return n
    out.push(n)
  }
  return out
}

function flag(v: Value): boolean | CellError {
  if (v === null || v === undefined) return false
  return toBool(v as Scalar)
}

/** Excel has no denormal numbers: what is below the smallest normal double shows as 0. */
const MIN_NORMAL = 2.2250738585072014e-308
const finite = (x: number): Scalar => (!Number.isFinite(x) ? NUM : Math.abs(x) < MIN_NORMAL ? 0 : x)

// ── Normal ──────────────────────────────────────────────────────────────────────────────────────────────────────

function normDist(a: Value[]): Scalar {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [x, mean, sd] = n
  if (sd <= 0) return NUM
  const z = (x - mean) / sd
  return finite(cumulative ? normCdf(z) : normPdf(z) / sd)
}
scalarFn("NORM.DIST", 4, 4, normDist)
scalarFn("NORMDIST", 4, 4, normDist)

function normInvFn(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, mean, sd] = n
  if (p <= 0 || p >= 1 || sd <= 0) return NUM
  return mean + sd * normInv(p)
}
scalarFn("NORM.INV", 3, 3, normInvFn)
scalarFn("NORMINV", 3, 3, normInvFn)

scalarFn("NORM.S.DIST", 2, 2, (a) => {
  const z = num(a[0])
  if (isError(z)) return z
  const cumulative = flag(a[1])
  if (isError(cumulative)) return cumulative
  return finite(cumulative ? normCdf(z) : normPdf(z))
})
scalarFn("NORMSDIST", 1, 1, (a) => {
  const z = num(a[0])
  return isError(z) ? z : finite(normCdf(z))
})
function normSInv(a: Value[]): Scalar {
  const p = num(a[0])
  if (isError(p)) return p
  return p <= 0 || p >= 1 ? NUM : normInv(p)
}
scalarFn("NORM.S.INV", 1, 1, normSInv)
scalarFn("NORMSINV", 1, 1, normSInv)

scalarFn("STANDARDIZE", 3, 3, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [x, mean, sd] = n
  return sd <= 0 ? NUM : (x - mean) / sd
})
scalarFn("GAUSS", 1, 1, (a) => {
  const z = num(a[0])
  if (isError(z)) return z
  // As Excel defines it, NORM.S.DIST(z, TRUE) − 0.5 (so small z loses digits the same way).
  return normCdf(z) - 0.5
})
scalarFn("PHI", 1, 1, (a) => {
  const x = num(a[0])
  return isError(x) ? x : normPdf(x)
})

// ── Lognormal ───────────────────────────────────────────────────────────────────────────────────────────────────

scalarFn("LOGNORM.DIST", 4, 4, (a) => {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [x, mean, sd] = n
  if (x <= 0 || sd <= 0) return NUM
  const z = (Math.log(x) - mean) / sd
  return cumulative ? normCdf(z) : normPdf(z) / (x * sd)
})
scalarFn("LOGNORMDIST", 3, 3, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [x, mean, sd] = n
  if (x <= 0 || sd <= 0) return NUM
  return normCdf((Math.log(x) - mean) / sd)
})
function logInv(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, mean, sd] = n
  if (p <= 0 || p >= 1 || sd <= 0) return NUM
  return Math.exp(mean + sd * normInv(p))
}
scalarFn("LOGNORM.INV", 3, 3, logInv)
scalarFn("LOGINV", 3, 3, logInv)

// ── Student's t ─────────────────────────────────────────────────────────────────────────────────────────────────

/**
 * P(T > t) for t ≥ 0. Whole degrees of freedom (as Excel truncates them) use the finite series in θ = atan(t/√ν);
 * others the incomplete beta (with 1 − x passed exactly, and the complement when t² is small next to df).
 */
function tUpper(t: number, df: number): number {
  if (Number.isInteger(df) && df <= 10000) {
    // The series gives P(|T| < t); its complement keeps its digits while the tail is not small.
    const tail = (1 - tCentral(t, df)) / 2
    if (tail >= 0.05) return tail
  }
  return tUpperBeta(t, df)
}
/** P(|T| < t) for whole ν (Abramowitz and Stegun 26.7.3 and 26.7.4). */
function tCentral(t: number, df: number): number {
  const r = Math.sqrt(df + t * t)
  const sin = t / r
  const cos = Math.sqrt(df) / r
  const cos2 = cos * cos
  if (df % 2 === 1) {
    const theta = Math.atan(t / Math.sqrt(df))
    if (df === 1) return (2 / Math.PI) * theta
    let term = cos
    let sum = cos
    for (let k = 2; k <= df - 2; k += 2) {
      term *= (cos2 * k) / (k + 1)
      sum += term
    }
    return (2 / Math.PI) * (theta + sin * sum)
  }
  let term = 1
  let sum = 1
  for (let k = 2; k <= df - 2; k += 2) {
    term *= (cos2 * (k - 1)) / k
    sum += term
  }
  return sin * sum
}
function tUpperBeta(t: number, df: number): number {
  const t2 = t * t
  const x = df / (df + t2)
  const y = t2 / (df + t2)
  return df > 2 * t2 ? 0.5 * betaIc(y, 0.5, df / 2, x) : 0.5 * betaI(x, df / 2, 0.5, y)
}
function tCdf(t: number, df: number): number {
  const upper = tUpper(Math.abs(t), df)
  return t >= 0 ? 1 - upper : upper
}
function tPdf(t: number, df: number): number {
  return Math.exp(
    logGamma((df + 1) / 2) -
      logGamma(df / 2) -
      0.5 * Math.log(df * Math.PI) -
      ((df + 1) / 2) * Math.log1p((t * t) / df),
  )
}
/** The t with P(T ≤ t) = p. */
function tInv(p: number, df: number): number {
  if (p === 0.5) return 0
  // Solve on the smaller tail: P(T > t) = q for t > 0.
  const q = p < 0.5 ? p : 1 - p
  const guess = Math.abs(normInv(q))
  const t = solveIncreasing(
    (x) => -tUpper(x, df),
    (x) => tPdf(x, df),
    -q,
    0,
    Infinity,
    guess,
  )
  return p < 0.5 ? -t : t
}

scalarFn("T.DIST", 3, 3, (a) => {
  const n = numbers(a, 0, 2)
  if (isError(n)) return n
  const cumulative = flag(a[2])
  if (isError(cumulative)) return cumulative
  const [x, df] = [n[0], Math.trunc(n[1])]
  if (df < 1) return NUM
  return cumulative ? tCdf(x, df) : tPdf(x, df)
})
scalarFn("T.DIST.2T", 2, 2, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [x, df] = [n[0], Math.trunc(n[1])]
  if (x < 0 || df < 1) return NUM
  return Math.min(1, 2 * tUpper(x, df))
})
scalarFn("T.DIST.RT", 2, 2, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [x, df] = [n[0], Math.trunc(n[1])]
  if (df < 1) return NUM
  return x >= 0 ? tUpper(x, df) : 1 - tUpper(-x, df)
})
scalarFn("TDIST", 3, 3, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [x, df, tails] = [n[0], Math.trunc(n[1]), Math.trunc(n[2])]
  if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) return NUM
  return tails * tUpper(x, df)
})
scalarFn("T.INV", 2, 2, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, df] = [n[0], Math.trunc(n[1])]
  if (p <= 0 || p >= 1 || df < 1) return NUM
  return tInv(p, df)
})
function tInv2(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, df] = [n[0], Math.trunc(n[1])]
  if (p <= 0 || p > 1 || df < 1) return NUM
  return Math.abs(tInv(p / 2, df))
}
scalarFn("T.INV.2T", 2, 2, tInv2)
scalarFn("TINV", 2, 2, tInv2)

// ── Chi-squared ─────────────────────────────────────────────────────────────────────────────────────────────────

function chisqPdf(x: number, k: number): number {
  if (x === 0) return k === 2 ? 0.5 : k < 2 ? Infinity : 0
  return gammaPrefix(k / 2, x / 2) / x
}
function chisqInv(p: number, k: number, upper: boolean): number {
  if (upper ? p === 1 : p === 0) return 0
  const f = upper ? (x: number) => -gammaQ(k / 2, x / 2) : (x: number) => gammaP(k / 2, x / 2)
  return solveIncreasing(f, (x) => chisqPdf(x, k), upper ? -p : p, 0, Infinity, k)
}
scalarFn("CHISQ.DIST", 3, 3, (a) => {
  const n = numbers(a, 0, 2)
  if (isError(n)) return n
  const cumulative = flag(a[2])
  if (isError(cumulative)) return cumulative
  const [x, k] = [n[0], Math.trunc(n[1])]
  if (x < 0 || k < 1 || k > 1e10) return NUM
  return cumulative ? gammaP(k / 2, x / 2) : finite(chisqPdf(x, k))
})
function chisqRt(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [x, k] = [n[0], Math.trunc(n[1])]
  if (x < 0 || k < 1 || k > 1e10) return NUM
  return gammaQ(k / 2, x / 2)
}
scalarFn("CHISQ.DIST.RT", 2, 2, chisqRt)
scalarFn("CHIDIST", 2, 2, chisqRt)
scalarFn("CHISQ.INV", 2, 2, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, k] = [n[0], Math.trunc(n[1])]
  if (p < 0 || p >= 1 || k < 1 || k > 1e10) return NUM
  return chisqInv(p, k, false)
})
function chisqInvRt(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, k] = [n[0], Math.trunc(n[1])]
  if (p <= 0 || p > 1 || k < 1 || k > 1e10) return NUM
  return chisqInv(p, k, true)
}
scalarFn("CHISQ.INV.RT", 2, 2, chisqInvRt)
scalarFn("CHIINV", 2, 2, chisqInvRt)

// ── F ───────────────────────────────────────────────────────────────────────────────────────────────────────────

function fCdf(x: number, d1: number, d2: number): number {
  const den = d1 * x + d2
  return betaI((d1 * x) / den, d1 / 2, d2 / 2, d2 / den)
}
function fUpper(x: number, d1: number, d2: number): number {
  const den = d1 * x + d2
  return betaI(d2 / den, d2 / 2, d1 / 2, (d1 * x) / den)
}
function fPdf(x: number, d1: number, d2: number): number {
  if (x === 0) return d1 === 2 ? 1 : d1 < 2 ? Infinity : 0
  const den = d1 * x + d2
  return betaPrefix((d1 * x) / den, d2 / den, d1 / 2, d2 / 2) / x
}
function fArgs(a: Value[], count: number): number[] | CellError {
  const n = numbers(a, 0, count)
  if (isError(n)) return n
  n[1] = Math.trunc(n[1])
  n[2] = Math.trunc(n[2])
  if (n[1] < 1 || n[2] < 1 || n[1] >= 1e10 || n[2] >= 1e10) return NUM
  return n
}
scalarFn("F.DIST", 4, 4, (a) => {
  const n = fArgs(a, 3)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [x, d1, d2] = n
  if (x < 0) return NUM
  return cumulative ? fCdf(x, d1, d2) : finite(fPdf(x, d1, d2))
})
function fRt(a: Value[]): Scalar {
  const n = fArgs(a, 3)
  if (isError(n)) return n
  const [x, d1, d2] = n
  return x < 0 ? NUM : fUpper(x, d1, d2)
}
scalarFn("F.DIST.RT", 3, 3, fRt)
scalarFn("FDIST", 3, 3, fRt)
function fInv(p: number, d1: number, d2: number, upper: boolean): number {
  if (upper ? p === 1 : p === 0) return 0
  const f = upper ? (x: number) => -fUpper(x, d1, d2) : (x: number) => fCdf(x, d1, d2)
  return solveIncreasing(f, (x) => fPdf(x, d1, d2), upper ? -p : p, 0, Infinity, 1)
}
scalarFn("F.INV", 3, 3, (a) => {
  const n = fArgs(a, 3)
  if (isError(n)) return n
  const [p, d1, d2] = n
  if (p < 0 || p >= 1) return NUM
  return fInv(p, d1, d2, false)
})
function fInvRt(a: Value[]): Scalar {
  const n = fArgs(a, 3)
  if (isError(n)) return n
  const [p, d1, d2] = n
  if (p <= 0 || p > 1) return NUM
  return fInv(p, d1, d2, true)
}
scalarFn("F.INV.RT", 3, 3, fInvRt)
scalarFn("FINV", 3, 3, fInvRt)

// ── Beta ────────────────────────────────────────────────────────────────────────────────────────────────────────

function betaPdf(x: number, a: number, b: number): number {
  if (x === 0) return a === 1 ? Math.exp(-logBeta(a, b)) : a < 1 ? Infinity : 0
  if (x === 1) return b === 1 ? Math.exp(-logBeta(a, b)) : b < 1 ? Infinity : 0
  return betaPrefix(x, 1 - x, a, b) / (x * (1 - x))
}
/** x, α, β and the bounds A and B (0 and 1 when left out). */
function betaArgs(
  a: Value[],
  from: number,
): { x: number; alpha: number; beta: number; lo: number; hi: number } | CellError {
  const head = numbers(a, 0, 3)
  if (isError(head)) return head
  const lo = a[from] == null ? 0 : num(a[from])
  const hi = a[from + 1] == null ? 1 : num(a[from + 1])
  if (isError(lo)) return lo
  if (isError(hi)) return hi
  const [x, alpha, beta] = head
  if (alpha <= 0 || beta <= 0 || lo === hi) return NUM
  return { x, alpha, beta, lo, hi }
}
scalarFn("BETA.DIST", 4, 6, (a) => {
  const b = betaArgs(a, 4)
  if (isError(b)) return b
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  if (b.x < b.lo || b.x > b.hi) return NUM
  const t = (b.x - b.lo) / (b.hi - b.lo)
  return cumulative ? betaI(t, b.alpha, b.beta) : finite(betaPdf(t, b.alpha, b.beta) / (b.hi - b.lo))
})
scalarFn("BETADIST", 3, 5, (a) => {
  const b = betaArgs(a, 3)
  if (isError(b)) return b
  if (b.x < b.lo || b.x > b.hi) return NUM
  return betaI((b.x - b.lo) / (b.hi - b.lo), b.alpha, b.beta)
})
function betaInv(a: Value[]): Scalar {
  const b = betaArgs(a, 3)
  if (isError(b)) return b
  const p = b.x
  if (p <= 0 || p > 1) return NUM
  const t =
    p === 1
      ? 1
      : solveIncreasing(
          (x) => betaI(x, b.alpha, b.beta),
          (x) => betaPdf(x, b.alpha, b.beta),
          p,
          0,
          1,
          b.alpha / (b.alpha + b.beta),
        )
  return b.lo + t * (b.hi - b.lo)
}
scalarFn("BETA.INV", 3, 5, betaInv)
scalarFn("BETAINV", 3, 5, betaInv)

// ── Gamma, exponential, Weibull ─────────────────────────────────────────────────────────────────────────────────

function gammaPdf(x: number, alpha: number, beta: number): number {
  if (x === 0) return alpha === 1 ? 1 / beta : alpha < 1 ? Infinity : 0
  return gammaPrefix(alpha, x / beta) / x
}
function gammaDist(a: Value[]): Scalar {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [x, alpha, beta] = n
  if (x < 0 || alpha <= 0 || beta <= 0) return NUM
  return cumulative ? gammaP(alpha, x / beta) : finite(gammaPdf(x, alpha, beta))
}
scalarFn("GAMMA.DIST", 4, 4, gammaDist)
scalarFn("GAMMADIST", 4, 4, gammaDist)
function gammaInv(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [p, alpha, beta] = n
  if (p < 0 || p >= 1 || alpha <= 0 || beta <= 0) return NUM
  if (p === 0) return 0
  return solveIncreasing(
    (x) => gammaP(alpha, x / beta),
    (x) => gammaPdf(x, alpha, beta),
    p,
    0,
    Infinity,
    alpha * beta,
  )
}
scalarFn("GAMMA.INV", 3, 3, gammaInv)
scalarFn("GAMMAINV", 3, 3, gammaInv)
scalarFn("GAMMA", 1, 1, (a) => {
  const x = num(a[0])
  if (isError(x)) return x
  if (x <= 0 && Number.isInteger(x)) return NUM
  return finite(gamma(x))
})
function gammaLn(a: Value[]): Scalar {
  const x = num(a[0])
  if (isError(x)) return x
  return x <= 0 ? NUM : logGamma(x)
}
scalarFn("GAMMALN", 1, 1, gammaLn)
scalarFn("GAMMALN.PRECISE", 1, 1, gammaLn)

function exponDist(a: Value[]): Scalar {
  const n = numbers(a, 0, 2)
  if (isError(n)) return n
  const cumulative = flag(a[2])
  if (isError(cumulative)) return cumulative
  const [x, lambda] = n
  if (x < 0 || lambda <= 0) return NUM
  return cumulative ? -Math.expm1(-lambda * x) : lambda * Math.exp(-lambda * x)
}
scalarFn("EXPON.DIST", 3, 3, exponDist)
scalarFn("EXPONDIST", 3, 3, exponDist)

function weibull(a: Value[]): Scalar {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [x, alpha, beta] = n
  if (x < 0 || alpha <= 0 || beta <= 0) return NUM
  const t = Math.pow(x / beta, alpha)
  // The density as Excel writes it: α/β^α · x^(α−1) · e^(−(x/β)^α).
  return cumulative ? -Math.expm1(-t) : (alpha / Math.pow(beta, alpha)) * Math.pow(x, alpha - 1) * Math.exp(-t)
}
scalarFn("WEIBULL.DIST", 4, 4, weibull)
scalarFn("WEIBULL", 4, 4, weibull)

// ── Discrete ────────────────────────────────────────────────────────────────────────────────────────────────────

/** ln C(n, k). */
function logChoose(n: number, k: number): number {
  return logGamma(n + 1) - logGamma(k + 1) - logGamma(n - k + 1)
}

/** C(n, k) as a float: exact while it fits, by the multiplicative formula. */
function choose(n: number, k: number): number {
  if (k < 0 || k > n) return 0
  k = Math.min(k, n - k)
  if (n > 1000) return Math.exp(logChoose(n, k))
  let r = 1
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i
  return r
}

function binomPmf(k: number, n: number, p: number): number {
  if (p === 0) return k === 0 ? 1 : 0
  if (p === 1) return k === n ? 1 : 0
  const c = choose(n, k)
  if (Number.isFinite(c) && c > 0 && n <= 1000) return c * Math.pow(p, k) * Math.pow(1 - p, n - k)
  return Math.exp(logChoose(n, k) + k * Math.log(p) + (n - k) * Math.log1p(-p))
}
function binomCdf(k: number, n: number, p: number): number {
  if (k >= n) return 1
  let s = 0
  for (let i = 0; i <= k; i++) s += binomPmf(i, n, p)
  return Math.min(1, s)
}

function binomDist(a: Value[]): Scalar {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [k, trials, p] = [Math.trunc(n[0]), Math.trunc(n[1]), n[2]]
  if (k < 0 || k > trials || p < 0 || p > 1) return NUM
  return cumulative ? binomCdf(k, trials, p) : binomPmf(k, trials, p)
}
scalarFn("BINOM.DIST", 4, 4, binomDist)
scalarFn("BINOMDIST", 4, 4, binomDist)
scalarFn("BINOM.DIST.RANGE", 3, 4, (a) => {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const [trials, p, s] = [Math.trunc(n[0]), n[1], Math.trunc(n[2])]
  const s2r = a[3] == null ? s : int(a[3])
  if (isError(s2r)) return s2r
  const s2 = s2r
  if (trials < 0 || p < 0 || p > 1 || s < 0 || s > trials || s2 < s || s2 > trials) return NUM
  let sum = 0
  for (let i = s; i <= s2; i++) sum += binomPmf(i, trials, p)
  return Math.min(1, sum)
})
function binomInv(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [trials, p, alpha] = [Math.trunc(n[0]), n[1], n[2]]
  if (trials < 0 || p < 0 || p > 1 || alpha <= 0 || alpha > 1) return NUM
  let s = 0
  for (let k = 0; k <= trials; k++) {
    s += binomPmf(k, trials, p)
    if (s >= alpha * (1 - 1e-15)) return k
  }
  return trials
}
scalarFn("BINOM.INV", 3, 3, binomInv)
scalarFn("CRITBINOM", 3, 3, binomInv)

function poissonPmf(k: number, mean: number): number {
  if (mean === 0) return k === 0 ? 1 : 0
  if (k <= 170 && mean < 700) {
    let r = Math.exp(-mean)
    for (let i = 1; i <= k; i++) r *= mean / i
    return r
  }
  return Math.exp(k * Math.log(mean) - mean - logGamma(k + 1))
}
function poisson(a: Value[]): Scalar {
  const n = numbers(a, 0, 2)
  if (isError(n)) return n
  const cumulative = flag(a[2])
  if (isError(cumulative)) return cumulative
  const [k, mean] = [Math.trunc(n[0]), n[1]]
  if (k < 0 || mean < 0) return NUM
  if (!cumulative) return poissonPmf(k, mean)
  // The upper incomplete gamma: Q(k + 1, mean).
  return mean === 0 ? 1 : gammaQ(k + 1, mean)
}
scalarFn("POISSON.DIST", 3, 3, poisson)
scalarFn("POISSON", 3, 3, poisson)

function hypgeomPmf(k: number, n: number, K: number, N: number): number {
  if (k < 0 || k > n || k > K || n - k > N - K) return 0
  const [a, b, c] = [choose(K, k), choose(N - K, n - k), choose(N, n)]
  if (c < 2 ** 53 && a * b < 2 ** 53) return (a * b) / c
  return Math.exp(logChoose(K, k) + logChoose(N - K, n - k) - logChoose(N, n))
}
function hypgeomArgs(a: Value[]): number[] | CellError {
  const n = numbers(a, 0, 4)
  if (isError(n)) return n
  const [k, sample, K, N] = n.map(Math.trunc)
  // Excel: a count of successes that can't happen is a probability of 0, not an error.
  if (k < 0 || sample <= 0 || sample > N || K <= 0 || K > N || N <= 0) return NUM
  return [k, sample, K, N]
}
scalarFn("HYPGEOM.DIST", 5, 5, (a) => {
  const n = hypgeomArgs(a)
  if (isError(n)) return n
  const cumulative = flag(a[4])
  if (isError(cumulative)) return cumulative
  const [k, sample, K, N] = n
  if (!cumulative) return hypgeomPmf(k, sample, K, N)
  let s = 0
  for (let i = Math.max(0, sample - N + K); i <= Math.min(k, sample, K); i++) s += hypgeomPmf(i, sample, K, N)
  return Math.min(1, s)
})
scalarFn("HYPGEOMDIST", 4, 4, (a) => {
  const n = hypgeomArgs(a)
  if (isError(n)) return n
  const [k, sample, K, N] = n
  return hypgeomPmf(k, sample, K, N)
})

function negbinomPmf(f: number, s: number, p: number): number {
  return choose(f + s - 1, s - 1) * Math.pow(p, s) * Math.pow(1 - p, f)
}
function negbinomArgs(a: Value[]): number[] | CellError {
  const n = numbers(a, 0, 3)
  if (isError(n)) return n
  const [f, s, p] = [Math.trunc(n[0]), Math.trunc(n[1]), n[2]]
  if (f < 0 || s < 1 || p < 0 || p > 1) return NUM
  return [f, s, p]
}
scalarFn("NEGBINOM.DIST", 4, 4, (a) => {
  const n = negbinomArgs(a)
  if (isError(n)) return n
  const cumulative = flag(a[3])
  if (isError(cumulative)) return cumulative
  const [f, s, p] = n
  if (!cumulative) return negbinomPmf(f, s, p)
  // The sum of the probabilities while they are few; the incomplete beta I_p(s, f + 1) beyond.
  if (f > 1000) return betaI(p, s, f + 1)
  let sum = 0
  for (let i = 0; i <= f; i++) sum += negbinomPmf(i, s, p)
  return Math.min(1, sum)
})
scalarFn("NEGBINOMDIST", 3, 3, (a) => {
  const n = negbinomArgs(a)
  if (isError(n)) return n
  const [f, s, p] = n
  return negbinomPmf(f, s, p)
})

// ── Confidence intervals, Fisher, PERMUTATIONA ──────────────────────────────────────────────────────────────────

function confidenceNorm(a: Value[]): Scalar {
  const n = numbers(a)
  if (isError(n)) return n
  const [alpha, sd, size] = [n[0], n[1], Math.trunc(n[2])]
  if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) return NUM
  return (-normInv(alpha / 2) * sd) / Math.sqrt(size)
}
scalarFn("CONFIDENCE.NORM", 3, 3, confidenceNorm)
scalarFn("CONFIDENCE", 3, 3, confidenceNorm)
scalarFn("CONFIDENCE.T", 3, 3, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [alpha, sd, size] = [n[0], n[1], Math.trunc(n[2])]
  if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) return NUM
  if (size === 1) return DIV0
  return (Math.abs(tInv(alpha / 2, size - 1)) * sd) / Math.sqrt(size)
})

scalarFn("FISHER", 1, 1, (a) => {
  const x = num(a[0])
  if (isError(x)) return x
  return x <= -1 || x >= 1 ? NUM : 0.5 * Math.log((1 + x) / (1 - x))
})
scalarFn("FISHERINV", 1, 1, (a) => {
  const y = num(a[0])
  if (isError(y)) return y
  return Math.tanh(y)
})
scalarFn("PERMUTATIONA", 2, 2, (a) => {
  const n = numbers(a)
  if (isError(n)) return n
  const [count, chosen] = [Math.trunc(n[0]), Math.trunc(n[1])]
  if (count < 0 || chosen < 0) return NUM
  return finite(Math.pow(count, chosen))
})

export { chisqInv, erfc, fUpper, gammaQ, normInv, tCdf, tInv, tUpper, VALUE }
