/**
 * The special functions behind the statistical distributions: the gamma function, the error function, the
 * regularized incomplete gamma and beta functions, and a safeguarded root finder for the inverse distributions. Each
 * aims at full double precision (checked against mpmath to 1e-14 relative), since Excel's results are compared to 15
 * significant digits.
 */

const EPS = 1e-16
const TINY = 1e-300
const SQRT_PI = Math.sqrt(Math.PI)
const LN_SQRT_2PI = 0.5 * Math.log(2 * Math.PI)

/** ln Γ(x) for x > 0 (Lanczos, g = 607/128, 15 terms, as in Boost and Python's math.lgamma). */
const LANCZOS_G = 607 / 128
const LANCZOS = [
  0.9999999999999971, 57.15623566586292, -59.59796035547549, 14.136097974741746, -0.4919138160976202,
  0.00003399464998481189, 0.00004652362892704858, -0.00009837447530487956, 0.0001580887032249125,
  -0.00021026444172410488, 0.00021743961811521265, -0.0001643181065367639, 0.00008441822398385275,
  -0.000026190838401581408, 0.0000036899182659531625,
]

function lanczosSum(x: number) {
  let s = LANCZOS[0]
  for (let i = 1; i < LANCZOS.length; i++) s += LANCZOS[i] / (x + i - 1)
  return s
}

/** Γ(x) for integers up to 170 exactly (as a product), from a table of factorials. */
const FACTORIALS: number[] = [1]
for (let i = 1; i <= 170; i++) FACTORIALS[i] = FACTORIALS[i - 1] * i

export function logGamma(x: number): number {
  if (x <= 0) return NaN
  if (Number.isInteger(x) && x <= 171) return Math.log(FACTORIALS[x - 1])
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - logGamma(1 - x)
  const t = x + LANCZOS_G - 0.5
  return LN_SQRT_2PI + Math.log(lanczosSum(x)) + (x - 0.5) * Math.log(t) - t
}

/** Γ(x), with the reflection formula for negative non-integers. */
export function gamma(x: number): number {
  if (Number.isInteger(x)) return x <= 0 ? NaN : x <= 171 ? FACTORIALS[x - 1] : Infinity
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x))
  if (x > 171.7) return Infinity
  const t = x + LANCZOS_G - 0.5
  // Split the power to keep it finite for x up to 171.
  const half = Math.pow(t, (x - 0.5) / 2)
  return Math.sqrt(2 * Math.PI) * lanczosSum(x) * half * (half * Math.exp(-t))
}

/** erf(x): a series of positive terms below 2.5, 1 − erfc(x) above. */
export function erf(x: number): number {
  if (x < 0) return -erf(-x)
  if (x >= 2.5) return 1 - erfc(x)
  // erf(x) = 2/√π · e^(−x²) · Σ 2ⁿ x^(2n+1) / (1·3·…·(2n+1)): no cancellation.
  let term = x
  let sum = x
  for (let n = 1; n < 500; n++) {
    term *= (2 * x * x) / (2 * n + 1)
    sum += term
    if (term < sum * EPS) break
  }
  return (2 / SQRT_PI) * Math.exp(-x * x) * sum
}

/** erfc(x) = 1 − erf(x), without losing digits for large x (a continued fraction). */
export function erfc(x: number): number {
  if (x < 0) return 2 - erfc(-x)
  if (x < 0.5) return 1 - erf(x)
  if (x > 27) return 0
  // erfc(x) = e^(−x²)/√π · 1/(x + (1/2)/(x + 1/(x + (3/2)/(x + …)))), modified Lentz.
  let f = x
  let c = x
  let d = 0
  for (let n = 1; n < 10000; n++) {
    const a = n / 2
    d = x + a * d
    d = d === 0 ? TINY : 1 / d
    c = x + a / c
    if (c === 0) c = TINY
    const delta = c * d
    f *= delta
    if (Math.abs(delta - 1) < EPS) break
  }
  return expMinusSquare(x) / SQRT_PI / f
}

/** e^(−x²) with x² kept to double-double precision (the rounding of x² alone costs x²·1e-16 otherwise). */
export function expMinusSquare(x: number): number {
  const [hi, lo] = twoProduct(x, x)
  return Math.exp(-hi) * (1 - lo)
}

/** a·b as an unevaluated sum hi + lo (Dekker's product with Veltkamp's split). */
function twoProduct(a: number, b: number): [number, number] {
  const hi = a * b
  const split = (v: number) => {
    const t = 134217729 * v
    const h = t - (t - v)
    return [h, v - h]
  }
  const [ah, al] = split(a)
  const [bh, bl] = split(b)
  const lo = ah * bh - hi + ah * bl + al * bh + al * bl
  return [hi, lo]
}

/** log(1 + d) − d, without cancellation for small d. */
function log1pmx(d: number): number {
  if (Math.abs(d) > 0.5) return Math.log1p(d) - d
  // −d²/2 + d³/3 − d⁴/4 + …
  let term = d
  let sum = 0
  for (let k = 2; k < 200; k++) {
    term *= -d
    const add = term / k
    sum += add
    if (Math.abs(add) < Math.abs(sum) * EPS) break
  }
  return sum
}

/**
 * z^a · e^(−z) / Γ(a), the factor in front of the incomplete gamma series and fraction, through the Lanczos sum so
 * that large a and z near a keep their digits (as Boost's regularised_gamma_prefix).
 */
export function gammaPrefix(a: number, z: number): number {
  if (z <= 0) return 0
  // Directly while nothing overflows: pow and exp are each good to an ulp, Γ is exact for integers.
  if (a < 150 && z < 700) {
    const direct = (Math.pow(z, a) * Math.exp(-z)) / gamma(a)
    if (Number.isFinite(direct) && direct > 1e-290) return direct
  }
  const t = a + LANCZOS_G - 0.5
  const d = (z - t) / t
  const power = a * log1pmx(d) - (LANCZOS_G - 0.5) * d
  return (Math.exp(power) * Math.sqrt(t)) / (Math.sqrt(2 * Math.PI) * lanczosSum(a))
}

/** x^a · y^b / B(a, b) with y = 1 − x, through the Lanczos sums (as Boost's ibeta_power_terms). */
export function betaPrefix(x: number, y: number, a: number, b: number): number {
  if (a + b < 150) {
    const direct = (Math.pow(x, a) * Math.pow(y, b) * gamma(a + b)) / (gamma(a) * gamma(b))
    if (Number.isFinite(direct) && direct > 1e-290) return direct
  }
  const g = LANCZOS_G - 0.5
  const [ag, bg, cg] = [a + g, b + g, a + b + g]
  const l1 = (x * cg - ag) / ag
  const l2 = (y * cg - bg) / bg
  const power = a * Math.log1p(l1) + b * Math.log1p(l2)
  const factor = (lanczosSum(a + b) / (lanczosSum(a) * lanczosSum(b))) * Math.sqrt((ag * bg) / cg) * Math.exp(g)
  return (factor * Math.exp(power)) / Math.sqrt(2 * Math.PI)
}

/** The standard normal distribution: Φ(z). */
export function normCdf(z: number): number {
  return 0.5 * erfc(-z / Math.SQRT2)
}

export function normPdf(z: number): number {
  return Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI)
}

/** Φ⁻¹(p): Acklam's approximation, then refined with Halley's method against erfc. */
export function normInv(p: number): number {
  if (p <= 0) return -Infinity
  if (p >= 1) return Infinity
  const a = [-39.69683028665376, 220.9460984245205, -275.928510446969, 138.357751867269, -30.66479806614716, 2.506628277459239]
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572]
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783]
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416]
  const low = 0.02425
  let x: number
  if (p < low) {
    const q = Math.sqrt(-2 * Math.log(p))
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  } else if (p <= 1 - low) {
    const q = p - 0.5
    const r = q * q
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p))
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  }
  for (let i = 0; i < 3; i++) {
    // Work on the smaller tail so the residual keeps its digits.
    const e = p < 0.5 ? 0.5 * erfc(-x / Math.SQRT2) - p : -(0.5 * erfc(x / Math.SQRT2) - (1 - p))
    const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2)
    const next = x - u / (1 + (x * u) / 2)
    if (next === x) break
    x = next
  }
  return x
} // prettier-ignore

/** The regularized lower incomplete gamma function P(a, x). */
export function gammaP(a: number, x: number): number {
  if (x <= 0) return 0
  if (x < a + 1) return gammaSeries(a, x)
  return 1 - gammaFraction(a, x)
}

/** The regularized upper incomplete gamma function Q(a, x) = 1 − P(a, x). */
export function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1
  if (x < a + 1) return 1 - gammaSeries(a, x)
  const finite = gammaQFinite(a, x)
  if (finite !== null) return finite
  return gammaFraction(a, x)
}

/**
 * Q(a, x) for a whole or half-whole a up to 100 as a finite sum of positive terms (the chi-squared and Poisson
 * cases): e^(−x)·Σ x^k/k! for whole a, erfc(√x) + e^(−x)·Σ x^(k+½)/Γ(k+3/2) for half-whole a.
 */
function gammaQFinite(a: number, x: number): number | null {
  if (a > 100 || x > 700 || !Number.isInteger(2 * a)) return null
  if (Number.isInteger(a)) {
    let term = 1
    let sum = 1
    for (let k = 1; k < a; k++) {
      term *= x / k
      sum += term
    }
    return sum * Math.exp(-x)
  }
  const root = Math.sqrt(x)
  let term = (root * 2) / SQRT_PI
  let sum = 0
  for (let k = 0; k + 0.5 < a; k++) {
    if (k > 0) term *= x / (k + 0.5)
    sum += term
  }
  return erfc(root) + sum * Math.exp(-x)
}

function gammaSeries(a: number, x: number): number {
  let sum = 1 / a
  let term = sum
  for (let n = 1; n < 100000; n++) {
    term *= x / (a + n)
    sum += term
    if (Math.abs(term) < Math.abs(sum) * EPS) break
  }
  return sum * gammaPrefix(a, x)
}

function gammaFraction(a: number, x: number): number {
  let b = x + 1 - a
  let c = 1 / TINY
  let d = 1 / b
  let h = d
  for (let i = 1; i < 100000; i++) {
    const an = -i * (i - a)
    b += 2
    d = an * d + b
    if (Math.abs(d) < TINY) d = TINY
    c = b + an / c
    if (Math.abs(c) < TINY) c = TINY
    d = 1 / d
    const delta = d * c
    h *= delta
    if (Math.abs(delta - 1) < EPS) break
  }
  return gammaPrefix(a, x) * h
}

/** ln B(a, b). */
export function logBeta(a: number, b: number): number {
  return logGamma(a) + logGamma(b) - logGamma(a + b)
}

/**
 * The regularized incomplete beta function I_x(a, b). `y` is 1 − x when the caller has it to more digits than 1 − x
 * would give (the t and F distributions do).
 */
export function betaI(x: number, a: number, b: number, y = 1 - x): number {
  if (x <= 0) return 0
  if (y <= 0) return 1
  const front = betaPrefix(x, y, a, b)
  // The continued fraction converges quickly for x < (a + 1)/(a + b + 2); use the symmetry otherwise.
  if (x < (a + 1) / (a + b + 2)) return (front * betaFraction(x, a, b)) / a
  return 1 - (front * betaFraction(y, b, a)) / b
}

/** 1 − I_x(a, b), without the cancellation of subtracting from 1. */
export function betaIc(x: number, a: number, b: number, y = 1 - x): number {
  if (x <= 0) return 1
  if (y <= 0) return 0
  const front = betaPrefix(x, y, a, b)
  if (x < (a + 1) / (a + b + 2)) return 1 - (front * betaFraction(x, a, b)) / a
  return (front * betaFraction(y, b, a)) / b
}

function betaFraction(x: number, a: number, b: number): number {
  const qab = a + b
  const qap = a + 1
  const qam = a - 1
  let c = 1
  let d = 1 - (qab * x) / qap
  if (Math.abs(d) < TINY) d = TINY
  d = 1 / d
  let h = d
  for (let m = 1; m < 100000; m++) {
    const m2 = 2 * m
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2))
    d = 1 + aa * d
    if (Math.abs(d) < TINY) d = TINY
    c = 1 + aa / c
    if (Math.abs(c) < TINY) c = TINY
    d = 1 / d
    h *= d * c
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2))
    d = 1 + aa * d
    if (Math.abs(d) < TINY) d = TINY
    c = 1 + aa / c
    if (Math.abs(c) < TINY) c = TINY
    d = 1 / d
    const delta = d * c
    h *= delta
    if (Math.abs(delta - 1) < EPS) break
  }
  return h
}

/**
 * The x in [lo, hi] where the increasing function f(x) equals target: bisection that switches to Newton steps (with
 * the derivative `df`) while they stay inside the bracket. `hi` may be Infinity: the bracket grows until it holds the
 * root.
 */
export function solveIncreasing(
  f: (x: number) => number,
  df: ((x: number) => number) | null,
  target: number,
  lo: number,
  hi: number,
  guess?: number,
): number {
  if (!Number.isFinite(hi)) {
    hi = Math.max(1, Math.abs(lo) * 2, guess ?? 1)
    for (let i = 0; i < 2000 && f(hi) < target; i++) {
      lo = hi
      hi *= 2
    }
  }
  if (!Number.isFinite(lo)) {
    lo = Math.min(-1, hi - 1, guess ?? -1)
    for (let i = 0; i < 2000 && f(lo) > target; i++) {
      hi = lo
      lo *= 2
    }
  }
  let x = guess !== undefined && guess > lo && guess < hi ? guess : (lo + hi) / 2
  for (let i = 0; i < 400; i++) {
    const fx = f(x) - target
    if (fx === 0) return x
    if (fx < 0) lo = x
    else hi = x
    let next = NaN
    if (df) {
      const slope = df(x)
      if (slope > 0 && Number.isFinite(slope)) next = x - fx / slope
    }
    if (!(next > lo && next < hi)) next = (lo + hi) / 2
    if (Math.abs(next - x) <= Math.abs(x) * 4e-16 || hi - lo <= Math.abs(x) * 4e-16) return next
    x = next
  }
  return x
}
