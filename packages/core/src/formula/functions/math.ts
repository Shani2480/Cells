import { isError, type CellError } from "../../values"
import { roundTo15 } from "../../values"
import { roundDecimal } from "../../rounding"
import { CALC, DIV0, NA, NUM, toNumber, VALUE } from "../coerce"
import { cleanSum, power } from "../evaluate"
import { isLambda, isMatrix, isRef, matrix, type Area, type Scalar, type Value } from "../types"
import { compileCriteria } from "./criteria"
import { arrayArg, collectNumbers, forEachScalar, int, num } from "./helpers"
import { define, scalarFn } from "./registry"

const n1 = (f: (x: number) => number | CellError) => (args: Value[]) => {
  const x = num(args[0])
  if (isError(x)) return x
  const r = f(x)
  return typeof r === "number" && !Number.isFinite(r) ? NUM : r
}

scalarFn("ROUND", 2, 2, (a) => {
  const x = num(a[0])
  const d = num(a[1])
  if (isError(x)) return x
  if (isError(d)) return d
  return roundDecimal(x, d)
})
scalarFn("ROUNDUP", 2, 2, (a) => {
  const x = num(a[0])
  const d = num(a[1])
  if (isError(x)) return x
  if (isError(d)) return d
  return roundDecimal(x, d, "up")
})
scalarFn("ROUNDDOWN", 2, 2, (a) => {
  const x = num(a[0])
  const d = num(a[1])
  if (isError(x)) return x
  if (isError(d)) return d
  return roundDecimal(x, d, "down")
})
scalarFn("TRUNC", 1, 2, (a) => {
  const x = num(a[0])
  const d = a.length > 1 ? num(a[1]) : 0
  if (isError(x)) return x
  if (isError(d)) return d
  return roundDecimal(x, d, "down")
})
scalarFn(
  "INT",
  1,
  1,
  // On the 15-digit value, like Excel: INT(4.35*100) is 435 although the product is 434.99999999999994.
  n1((x) => Math.floor(roundTo15(x))),
)
scalarFn("ABS", 1, 1, n1(Math.abs))
scalarFn(
  "SIGN",
  1,
  1,
  n1((x) => Math.sign(x)),
)
scalarFn(
  "SQRT",
  1,
  1,
  n1((x) => (x < 0 ? NUM : Math.sqrt(x))),
)
scalarFn(
  "SQRTPI",
  1,
  1,
  n1((x) => (x < 0 ? NUM : Math.sqrt(x * Math.PI))),
)
scalarFn("EXP", 1, 1, n1(Math.exp))
scalarFn(
  "LN",
  1,
  1,
  n1((x) => (x <= 0 ? NUM : Math.log(x))),
)
scalarFn(
  "LOG10",
  1,
  1,
  n1((x) => (x <= 0 ? NUM : Math.log10(x))),
)
scalarFn("LOG", 1, 2, (a) => {
  const x = num(a[0])
  const b = a.length > 1 && a[1] !== null ? num(a[1]) : 10
  if (isError(x)) return x
  if (isError(b)) return b
  if (x <= 0 || b <= 0) return NUM
  if (b === 1) return DIV0
  return Math.log(x) / Math.log(b)
})
scalarFn("POWER", 2, 2, (a) => {
  const x = num(a[0])
  const y = num(a[1])
  if (isError(x)) return x
  if (isError(y)) return y
  return power(x, y)
})
define("PI", { min: 0, max: 0, call: () => Math.PI })
scalarFn("MOD", 2, 2, (a) => {
  const x = num(a[0])
  const y = num(a[1])
  if (isError(x)) return x
  if (isError(y)) return y
  if (y === 0) return DIV0
  // C's fmod (MOD(10,0.1) is 0.0999…45 in Excel too), then the result takes the divisor's sign: MOD(-3, 2) is 1.
  const r = x % y
  return r !== 0 && r < 0 !== y < 0 ? r + y : r
})
scalarFn("QUOTIENT", 2, 2, (a) => {
  const x = num(a[0])
  const y = num(a[1])
  if (isError(x)) return x
  if (isError(y)) return y
  if (y === 0) return DIV0
  return Math.trunc(x / y)
})
scalarFn(
  "EVEN",
  1,
  1,
  n1((x) => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2),
)
scalarFn(
  "ODD",
  1,
  1,
  n1((x) => {
    const s = x < 0 ? -1 : 1
    let r = Math.ceil(Math.abs(x))
    if (r % 2 === 0) r += 1
    return s * r
  }),
)
scalarFn(
  "FACT",
  1,
  1,
  n1((x) => {
    if (x < 0) return NUM
    // From the top down, as Excel multiplies (FACT(170) differs in the last digits otherwise).
    let r = 1
    for (let i = Math.floor(x); i >= 2; i--) r *= i
    return r
  }),
)
scalarFn(
  "FACTDOUBLE",
  1,
  1,
  n1((x) => {
    if (x < -1) return NUM
    let r = 1
    for (let i = Math.floor(x); i > 1; i -= 2) r *= i
    return r
  }),
)
function combin(n: number, k: number) {
  if (k > n - k) k = n - k
  let r = 1
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i
  return Math.round(r)
}
scalarFn("COMBIN", 2, 2, (a) => {
  const n = int(a[0])
  const k = int(a[1])
  if (isError(n)) return n
  if (isError(k)) return k
  if (n < 0 || k < 0 || k > n) return NUM
  return combin(n, k)
})
scalarFn("COMBINA", 2, 2, (a) => {
  const n = int(a[0])
  const k = int(a[1])
  if (isError(n)) return n
  if (isError(k)) return k
  if (n < 0 || k < 0) return NUM
  return n === 0 && k === 0 ? 1 : combin(n + k - 1, k)
})
scalarFn("PERMUT", 2, 2, (a) => {
  const n = int(a[0])
  const k = int(a[1])
  if (isError(n)) return n
  if (isError(k)) return k
  if (n < 0 || k < 0 || k > n) return NUM
  let r = 1
  for (let i = 0; i < k; i++) r *= n - i
  return r
})

function gcd2(a: number, b: number): number {
  while (b) [a, b] = [b, a % b]
  return a
}
function integerList(args: Value[], ctx: Parameters<typeof collectNumbers>[1]): number[] | CellError {
  const ns = collectNumbers(args, ctx)
  if (isError(ns)) return ns
  if (ns.some((x) => x < 0)) return NUM
  return ns.map(Math.trunc)
}
define("GCD", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ns = integerList(args, ctx)
    return isError(ns) ? ns : ns.reduce(gcd2, 0)
  },
})
define("LCM", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ns = integerList(args, ctx)
    if (isError(ns)) return ns
    return ns.reduce((acc, x) => (acc === 0 || x === 0 ? 0 : (acc / gcd2(acc, x)) * x), 1)
  },
})

/** CEILING(x, significance): significance must share x's sign (except a negative x with a positive one). */
function ceilingTo(x: number, s: number, away: boolean) {
  if (s === 0) return 0
  const q = roundTo15(x / s)
  const r = away ? (q < 0 ? Math.floor(q) : Math.ceil(q)) : Math.ceil(q)
  return roundTo15(r * s)
}
scalarFn("CEILING", 2, 2, (a) => {
  const x = num(a[0])
  const s = num(a[1])
  if (isError(x)) return x
  if (isError(s)) return s
  if (x > 0 && s < 0) return NUM
  if (s === 0) return 0
  // Negative number and negative significance round away from zero.
  return x < 0 && s < 0 ? -ceilingTo(-x, -s, false) : ceilingTo(x, s, false)
})
scalarFn("CEILING.MATH", 1, 3, (a) => {
  const x = num(a[0])
  const s0 = a.length > 1 && a[1] !== null ? num(a[1]) : 1
  const mode = a.length > 2 && a[2] !== null ? num(a[2]) : 0
  if (isError(x)) return x
  if (isError(s0)) return s0
  if (isError(mode)) return mode
  const s = Math.abs(s0)
  if (s === 0) return 0
  if (x < 0 && mode !== 0) return -ceilingTo(-x, s, false)
  return ceilingTo(x, s, false)
})
scalarFn("CEILING.PRECISE", 1, 2, (a) => {
  const x = num(a[0])
  const s = a.length > 1 && a[1] !== null ? num(a[1]) : 1
  if (isError(x)) return x
  if (isError(s)) return s
  return s === 0 ? 0 : ceilingTo(x, Math.abs(s), false)
})
scalarFn("ISO.CEILING", 1, 2, (a) => {
  const x = num(a[0])
  const s = a.length > 1 && a[1] !== null ? num(a[1]) : 1
  if (isError(x)) return x
  if (isError(s)) return s
  return s === 0 ? 0 : ceilingTo(x, Math.abs(s), false)
})
function floorTo(x: number, s: number) {
  if (s === 0) return 0
  // The quotient on its 15-digit value, as INT: FLOOR(4.35*100,1) is 435.
  return roundTo15(Math.floor(roundTo15(x / s)) * s)
}
scalarFn("FLOOR", 2, 2, (a) => {
  const x = num(a[0])
  const s = num(a[1])
  if (isError(x)) return x
  if (isError(s)) return s
  if (x > 0 && s < 0) return NUM
  if (s === 0) return x === 0 ? 0 : DIV0
  return x < 0 && s < 0 ? -floorTo(-x, -s) : floorTo(x, s)
})
scalarFn("FLOOR.MATH", 1, 3, (a) => {
  const x = num(a[0])
  const s0 = a.length > 1 && a[1] !== null ? num(a[1]) : 1
  const mode = a.length > 2 && a[2] !== null ? num(a[2]) : 0
  if (isError(x)) return x
  if (isError(s0)) return s0
  if (isError(mode)) return mode
  const s = Math.abs(s0)
  if (s === 0) return 0
  if (x < 0 && mode !== 0) return -floorTo(-x, s)
  return floorTo(x, s)
})
scalarFn("FLOOR.PRECISE", 1, 2, (a) => {
  const x = num(a[0])
  const s = a.length > 1 && a[1] !== null ? num(a[1]) : 1
  if (isError(x)) return x
  if (isError(s)) return s
  return s === 0 ? 0 : floorTo(x, Math.abs(s))
})
scalarFn("MROUND", 2, 2, (a) => {
  const x = num(a[0])
  const m = num(a[1])
  if (isError(x)) return x
  if (isError(m)) return m
  if (m === 0) return 0
  if (Math.sign(x) * Math.sign(m) < 0) return NUM
  return roundTo15(roundDecimal(x / m, 0) * m)
})

// Trigonometry.
scalarFn("SIN", 1, 1, n1(Math.sin))
scalarFn("COS", 1, 1, n1(Math.cos))
scalarFn("TAN", 1, 1, n1(Math.tan))
scalarFn(
  "ASIN",
  1,
  1,
  n1((x) => (Math.abs(x) > 1 ? NUM : Math.asin(x))),
)
scalarFn(
  "ACOS",
  1,
  1,
  n1((x) => (Math.abs(x) > 1 ? NUM : Math.acos(x))),
)
scalarFn("ATAN", 1, 1, n1(Math.atan))
scalarFn("ATAN2", 2, 2, (a) => {
  const x = num(a[0])
  const y = num(a[1])
  if (isError(x)) return x
  if (isError(y)) return y
  if (x === 0 && y === 0) return DIV0
  return Math.atan2(y, x)
})
scalarFn("SINH", 1, 1, n1(Math.sinh))
scalarFn("COSH", 1, 1, n1(Math.cosh))
scalarFn("TANH", 1, 1, n1(Math.tanh))
scalarFn("ASINH", 1, 1, n1(Math.asinh))
scalarFn(
  "ACOSH",
  1,
  1,
  n1((x) => (x < 1 ? NUM : Math.acosh(x))),
)
scalarFn(
  "ATANH",
  1,
  1,
  n1((x) => (Math.abs(x) >= 1 ? NUM : Math.atanh(x))),
)
scalarFn(
  "COT",
  1,
  1,
  n1((x) => (x === 0 ? DIV0 : 1 / Math.tan(x))),
)
scalarFn(
  "CSC",
  1,
  1,
  n1((x) => (x === 0 ? DIV0 : 1 / Math.sin(x))),
)
scalarFn(
  "SEC",
  1,
  1,
  n1((x) => 1 / Math.cos(x)),
)
scalarFn(
  "DEGREES",
  1,
  1,
  n1((x) => (x * 180) / Math.PI),
)
scalarFn(
  "RADIANS",
  1,
  1,
  n1((x) => (x * Math.PI) / 180),
)

// Random numbers (volatile).
define("RAND", { min: 0, max: 0, volatile: true, call: (_a, ctx) => ctx.random() })
scalarFn(
  "RANDBETWEEN",
  2,
  2,
  (a, ctx) => {
    const lo = num(a[0])
    const hi = num(a[1])
    if (isError(lo)) return lo
    if (isError(hi)) return hi
    const l = Math.ceil(lo)
    const h = Math.floor(hi)
    if (l > h) return NUM
    return l + Math.floor(ctx.random() * (h - l + 1))
  },
  { volatile: true },
)
define("RANDARRAY", {
  min: 0,
  max: 5,
  volatile: true,
  params: ["scalar"],
  call: (a, ctx) => {
    const rows = a[0] == null ? 1 : int(a[0])
    const cols = a[1] == null ? 1 : int(a[1])
    const lo = a[2] == null ? 0 : num(a[2])
    const hi = a[3] == null ? 1 : num(a[3])
    const whole = a[4] == null ? false : a[4] === true || (typeof a[4] === "number" && a[4] !== 0)
    for (const v of [rows, cols, lo, hi]) if (isError(v)) return v
    const [r, c, l, h] = [rows, cols, lo, hi] as number[]
    if (r < 1 || c < 1 || l > h) return r === 0 || c === 0 ? CALC : VALUE
    const data: Scalar[] = []
    for (let i = 0; i < r * c; i++) {
      data.push(
        whole
          ? Math.ceil(l) + Math.floor(ctx.random() * (Math.floor(h) - Math.ceil(l) + 1))
          : l + ctx.random() * (h - l),
      )
    }
    return matrix(r, c, data)
  },
})

define("SEQUENCE", {
  min: 1,
  max: 4,
  params: ["scalar"],
  call: (a) => {
    const rows = int(a[0])
    const cols = a[1] == null ? 1 : int(a[1])
    const start = a[2] == null ? 1 : num(a[2])
    const step = a[3] == null ? 1 : num(a[3])
    for (const v of [rows, cols, start, step]) if (isError(v)) return v
    const [r, c, s, st] = [rows, cols, start, step] as number[]
    if (r < 0 || c < 0) return VALUE
    if (r === 0 || c === 0) return CALC
    if (r * c > 1_048_576 * 16_384) return NUM
    const data: Scalar[] = new Array(r * c)
    for (let i = 0; i < r * c; i++) data[i] = s + i * st
    return matrix(r, c, data)
  },
})

// ── Sums ────────────────────────────────────────────────────────────────────────────────────────────────────────

/** SUM adds left to right and cleans up cancellation in its last addition only (as Excel does). */
export function excelSum(ns: number[]): number {
  let total = 0
  for (let i = 0; i < ns.length; i++) {
    const next = total + ns[i]
    total = i === ns.length - 1 ? cleanSum(next, total, ns[i]) : next
  }
  return total
}

define("SUM", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ns = collectNumbers(args, ctx)
    if (isError(ns)) return ns
    const r = excelSum(ns)
    return Number.isFinite(r) ? r : NUM
  },
})
define("PRODUCT", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ns = collectNumbers(args, ctx)
    if (isError(ns)) return ns
    if (ns.length === 0) return 0
    return ns.reduce((p, x) => p * x, 1)
  },
})
define("SUMSQ", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ns = collectNumbers(args, ctx)
    return isError(ns) ? ns : ns.reduce((s, x) => s + x * x, 0)
  },
})

/** SUMPRODUCT: equal-sized arrays multiplied element by element; non-numbers count as 0. */
define("SUMPRODUCT", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    const ms = []
    for (const arg of args) {
      const m = arrayArg(arg, ctx)
      if (isError(m)) return m
      ms.push(m)
    }
    const { rows, cols } = ms[0]
    if (ms.some((m) => m.rows !== rows || m.cols !== cols)) return VALUE
    let total = 0
    for (let i = 0; i < rows * cols; i++) {
      let p = 1
      for (const m of ms) {
        const v = m.data[i]
        if (isError(v)) return v
        p *= typeof v === "number" ? v : 0
      }
      total += p
    }
    return total
  },
})

/** The cells of `range` paired with the cells of `sumRange` (same shape, anchored at its top-left). */
function pairedArea(range: Value, sumRange: Value | undefined) {
  const r = isRef(range) ? range.areas[0] : null
  if (!r) return null
  if (sumRange === undefined) return { criteria: r, values: r }
  const s = isRef(sumRange) ? sumRange.areas[0] : null
  if (!s) return null
  // Excel sizes the sum range like the criteria range, from the sum range's top-left cell.
  return {
    criteria: r,
    values: { sheet: s.sheet, r0: s.r0, c0: s.c0, r1: s.r0 + (r.r1 - r.r0), c1: s.c0 + (r.c1 - r.c0) },
  }
}

type Ctx = Parameters<typeof collectNumbers>[1]

/** Visits the cells of criteria ranges that meet all criteria, giving their offset inside the first range. */
export function matchingOffsets(pairs: { range: Value; criteria: Value }[], ctx: Ctx): number[] | CellError {
  const areas: Area[] = []
  for (const p of pairs) {
    if (!isRef(p.range) || p.range.areas.length !== 1) return VALUE
    areas.push(p.range.areas[0])
  }
  const first = areas[0]
  const rows = first.r1 - first.r0 + 1
  const cols = first.c1 - first.c0 + 1
  if (areas.some((a) => a.r1 - a.r0 + 1 !== rows || a.c1 - a.c0 + 1 !== cols)) return VALUE
  const tests = pairs.map((p) => compileCriteria(p.criteria as Scalar))
  // A blank cell can match some criteria ("", "<>x", "=" …): test whether it does, to avoid visiting blanks.
  const blankMatches = tests.map((t) => t(null))
  let candidates: Set<number> | null = null
  if (blankMatches.some((b) => !b)) {
    // Only non-blank cells of a range whose criteria rejects blanks can match: start from those.
    const k = blankMatches.indexOf(false)
    const a = areas[k]
    const used = ctx.usedArea(a)
    candidates = new Set()
    if (used) {
      const t = tests[k]
      ctx.values(used, (v, row, col) => {
        if (t(v)) candidates!.add((row - a.r0) * cols + (col - a.c0))
      })
    }
  }
  const out: number[] = []
  const check = (offset: number) => {
    const dr = Math.floor(offset / cols)
    const dc = offset % cols
    for (let k = 0; k < areas.length; k++) {
      const a = areas[k]
      if (!tests[k](ctx.cell(a.sheet, a.r0 + dr, a.c0 + dc))) return
    }
    out.push(offset)
  }
  if (candidates) {
    for (const offset of [...candidates].sort((x, y) => x - y)) check(offset)
  } else {
    // Every criteria matches blanks: limit to the used part of the sheet (beyond it everything is blank).
    const used = ctx.usedArea(first)
    const limitRows = used ? Math.min(rows, used.r1 - first.r0 + 1) : 0
    for (let offset = 0; offset < limitRows * cols; offset++) check(offset)
  }
  return out
}

function sumAt(area: { sheet: number; r0: number; c0: number; c1: number }, cols: number, offsets: number[], ctx: Ctx) {
  const ns: number[] = []
  for (const o of offsets) {
    const v = ctx.cell(area.sheet, area.r0 + Math.floor(o / cols), area.c0 + (o % cols))
    if (isError(v)) return v
    if (typeof v === "number") ns.push(v)
  }
  return ns
}

define("SUMIF", {
  min: 2,
  max: 3,
  params: ["any", "scalar", "any"],
  errors: true,
  call: (a, ctx) => {
    const p = pairedArea(a[0], a[2] ?? undefined)
    if (!p) return VALUE
    const offsets = matchingOffsets([{ range: { kind: "ref", areas: [p.criteria] }, criteria: a[1] }], ctx)
    if (isError(offsets)) return offsets
    const ns = sumAt(p.values, p.criteria.c1 - p.criteria.c0 + 1, offsets, ctx)
    return isError(ns) ? ns : excelSum(ns)
  },
})
define("SUMIFS", {
  min: 3,
  max: 255,
  errors: true,
  call: (a, ctx) => {
    if (a.length % 2 === 0) return VALUE
    const pairs = []
    for (let i = 1; i < a.length; i += 2) pairs.push({ range: a[i], criteria: scalarOf(a[i + 1], ctx) })
    const offsets = matchingOffsets(pairs, ctx)
    if (isError(offsets)) return offsets
    const sumArea = isRef(a[0]) ? a[0].areas[0] : null
    const first = isRef(a[1]) ? a[1].areas[0] : null
    if (!sumArea || !first) return VALUE
    if (sumArea.r1 - sumArea.r0 !== first.r1 - first.r0 || sumArea.c1 - sumArea.c0 !== first.c1 - first.c0) return VALUE
    const ns = sumAt(sumArea, first.c1 - first.c0 + 1, offsets, ctx)
    return isError(ns) ? ns : excelSum(ns)
  },
})

export function scalarOf(v: Value, ctx: Ctx): Scalar {
  if (isRef(v)) {
    const a = v.areas[0]
    return ctx.cell(a.sheet, a.r0, a.c0)
  }
  if (isMatrix(v)) return v.data[0]
  if (isLambda(v)) return VALUE
  return v
}

/** SUBTOTAL and AGGREGATE skip cells holding SUBTOTAL/AGGREGATE formulas (and hidden rows for 101–111). */
export const SUBTOTAL_SKIP = /^=\s*(?:SUBTOTAL|AGGREGATE)\s*\(/i

define("SUMX2MY2", {
  min: 2,
  max: 2,
  call: (a, ctx) => pairwise(a, ctx, (x, y) => x * x - y * y),
})
define("SUMX2PY2", {
  min: 2,
  max: 2,
  call: (a, ctx) => pairwise(a, ctx, (x, y) => x * x + y * y),
})
define("SUMXMY2", {
  min: 2,
  max: 2,
  call: (a, ctx) => pairwise(a, ctx, (x, y) => (x - y) ** 2),
})
function pairwise(a: Value[], ctx: Ctx, f: (x: number, y: number) => number) {
  const x = arrayArg(a[0], ctx)
  const y = arrayArg(a[1], ctx)
  if (isError(x)) return x
  if (isError(y)) return y
  if (x.data.length !== y.data.length) return NA
  let total = 0
  for (let i = 0; i < x.data.length; i++) {
    const p = x.data[i]
    const q = y.data[i]
    if (isError(p)) return p
    if (isError(q)) return q
    if (typeof p === "number" && typeof q === "number") total += f(p, q)
  }
  return total
}

// ── Matrices and number systems ──────────────────────────────────────────────────────────────────────────────

define("MMULT", {
  min: 2,
  max: 2,
  call: (a, ctx) => {
    const x = arrayArg(a[0], ctx)
    if (isError(x)) return x
    const y = arrayArg(a[1], ctx)
    if (isError(y)) return y
    if (x.cols !== y.rows) return VALUE
    // Every cell must be a number: blanks, text and logicals give #VALUE!.
    for (const v of [...x.data, ...y.data]) {
      if (isError(v)) return v
      if (typeof v !== "number") return VALUE
    }
    const data: number[] = new Array(x.rows * y.cols)
    for (let r = 0; r < x.rows; r++) {
      for (let c = 0; c < y.cols; c++) {
        let sum = 0
        for (let k = 0; k < x.cols; k++) sum += (x.data[r * x.cols + k] as number) * (y.data[k * y.cols + c] as number)
        data[r * y.cols + c] = sum
      }
    }
    return matrix(x.rows, y.cols, data)
  },
})

scalarFn("BASE", 2, 3, (a) => {
  const n = int(a[0])
  const radix = int(a[1])
  const min = a.length > 2 && a[2] !== null ? int(a[2]) : 0
  if (isError(n)) return n
  if (isError(radix)) return radix
  if (isError(min)) return min
  if (n < 0 || n >= 2 ** 53 || radix < 2 || radix > 36 || min < 0 || min > 255) return NUM
  return n.toString(radix).toUpperCase().padStart(min, "0")
})

scalarFn("DECIMAL", 2, 2, (a) => {
  const text = a[0] === null ? "" : String(a[0])
  const radix = int(a[1])
  if (isError(a[0] as Scalar)) return a[0] as CellError
  if (isError(radix)) return radix
  if (radix < 2 || radix > 36 || text.length > 255) return NUM
  let n = 0
  for (const ch of text.toUpperCase()) {
    const d = parseInt(ch, 36)
    if (Number.isNaN(d) || d >= radix) return NUM
    n = n * radix + d
  }
  return n
})

const ROMAN_DIGITS: [number, string][] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
]

/** ROMAN in the classic form (form 0). The concise forms 1-4 are not implemented yet and give the classic form. */
scalarFn("ROMAN", 1, 2, (a) => {
  let n = int(a[0])
  if (isError(n)) return n
  if (n < 0 || n > 3999) return VALUE
  let out = ""
  for (const [value, letters] of ROMAN_DIGITS) {
    while (n >= value) {
      out += letters
      n -= value
    }
  }
  return out
})

scalarFn("ARABIC", 1, 1, (a) => {
  if (isError(a[0] as Scalar)) return a[0] as CellError
  let text = (a[0] === null ? "" : String(a[0])).trim().toUpperCase()
  const negative = text.startsWith("-")
  if (negative) text = text.slice(1)
  if (text.length > 255) return VALUE
  const values: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }
  let total = 0
  for (let i = 0; i < text.length; i++) {
    const v = values[text[i]]
    if (v === undefined) return VALUE
    const next = values[text[i + 1]] ?? 0
    total += v < next ? -v : v
  }
  return negative ? -total : total
})

export { forEachScalar, roundDecimal, toNumber }
