import { isError, type CellError } from "../../values"
import { compareScalars, DIV0, NA, NUM, toNumber, VALUE } from "../coerce"
import { isRef, type Scalar, type Value } from "../types"
import { compileCriteria } from "./criteria"
import { arrayArg, collectNumbers, forEachScalar, int, num } from "./helpers"
import { excelSum, matchingOffsets, scalarOf } from "./math"
import { define } from "./registry"

type Ctx = Parameters<typeof collectNumbers>[1]

const numbersFn = (name: string, f: (ns: number[]) => Scalar, options: { logicalsInRefs?: boolean } = {}) =>
  define(name, {
    min: 1,
    max: 255,
    call: (args, ctx) => {
      const ns = collectNumbers(args, ctx, options)
      return isError(ns) ? ns : f(ns)
    },
  })

const average = (ns: number[]) => (ns.length ? excelSum(ns) / ns.length : DIV0)
numbersFn("AVERAGE", average)
numbersFn("AVERAGEA", average, { logicalsInRefs: true })
numbersFn("MIN", (ns) => (ns.length ? Math.min(...ns) : 0))
numbersFn("MAX", (ns) => (ns.length ? Math.max(...ns) : 0))
numbersFn("MINA", (ns) => (ns.length ? Math.min(...ns) : 0), { logicalsInRefs: true })
numbersFn("MAXA", (ns) => (ns.length ? Math.max(...ns) : 0), { logicalsInRefs: true })

export function median(ns: number[]): Scalar {
  if (!ns.length) return NUM
  const s = [...ns].sort((a, b) => a - b)
  const mid = s.length >> 1
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2
}
numbersFn("MEDIAN", median)

export function variance(ns: number[], sample: boolean): Scalar {
  const n = ns.length
  if (n === 0 || (sample && n === 1)) return DIV0
  // Excel's own formula, Σx² − (Σx)²/n (checked bit for bit: STDEV(0.1,0.2,0.3) is 0.09999999999999994).
  let s1 = 0
  let s2 = 0
  for (const x of ns) {
    s1 += x
    s2 += x * x
  }
  const ss = Math.max(0, s2 - (s1 * s1) / n)
  return ss / (sample ? n - 1 : n)
}
const sd = (v: Scalar) => (typeof v === "number" ? Math.sqrt(v) : v)
for (const [name, sample, a] of [
  ["VAR.S", true, false],
  ["VAR", true, false],
  ["VAR.P", false, false],
  ["VARP", false, false],
  ["VARA", true, true],
  ["VARPA", false, true],
] as const) {
  numbersFn(name, (ns) => variance(ns, sample), { logicalsInRefs: a })
}
for (const [name, sample, a] of [
  ["STDEV.S", true, false],
  ["STDEV", true, false],
  ["STDEV.P", false, false],
  ["STDEVP", false, false],
  ["STDEVA", true, true],
  ["STDEVPA", false, true],
] as const) {
  numbersFn(name, (ns) => sd(variance(ns, sample)), { logicalsInRefs: a })
}
numbersFn("DEVSQ", (ns) => {
  const mean = ns.reduce((a, b) => a + b, 0) / (ns.length || 1)
  return ns.reduce((a, x) => a + (x - mean) ** 2, 0)
})
numbersFn("AVEDEV", (ns) => {
  if (!ns.length) return NUM
  const mean = ns.reduce((a, b) => a + b, 0) / ns.length
  return ns.reduce((a, x) => a + Math.abs(x - mean), 0) / ns.length
})
numbersFn("GEOMEAN", (ns) => {
  if (!ns.length || ns.some((x) => x <= 0)) return NUM
  return Math.exp(ns.reduce((a, x) => a + Math.log(x), 0) / ns.length)
})
numbersFn("HARMEAN", (ns) => {
  if (!ns.length || ns.some((x) => x <= 0)) return NUM
  return ns.length / ns.reduce((a, x) => a + 1 / x, 0)
})

function mode(ns: number[]): Scalar {
  const counts = new Map<number, number>()
  let best: number | null = null
  let bestCount = 1
  for (const x of ns) {
    const c = (counts.get(x) ?? 0) + 1
    counts.set(x, c)
    if (c > bestCount) {
      bestCount = c
      best = x
    }
  }
  // Ties go to the value that appeared first.
  if (best !== null) {
    for (const x of ns) if (counts.get(x) === bestCount) return x
  }
  return NA
}
numbersFn("MODE", mode)
numbersFn("MODE.SNGL", mode)

// ── Counting ────────────────────────────────────────────────────────────────────────────────────────────────────

define("COUNT", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    let n = 0
    forEachScalar(args, ctx, (v, direct) => {
      if (typeof v === "number") n++
      else if (direct && (typeof v === "boolean" || (typeof v === "string" && !isError(toNumber(v))))) n++
    })
    return n
  },
})
define("COUNTA", {
  min: 1,
  max: 255,
  call: (args, ctx) => {
    let n = 0
    forEachScalar(args, ctx, () => {
      n++
    })
    // A typed empty argument (COUNTA(A1,)) counts as a value.
    return n + args.filter((a) => a === null).length
  },
})
define("COUNTBLANK", {
  min: 1,
  max: 1,
  call: (args, ctx) => {
    const v = args[0]
    if (!isRef(v)) return VALUE
    const a = v.areas[0]
    let filled = 0
    ctx.values(a, (x) => {
      if (x !== "") filled++
    })
    return (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) - filled
  },
})
define("COUNTIF", {
  min: 2,
  max: 2,
  params: ["any", "scalar"],
  errors: true,
  call: (a, ctx) => {
    const offsets = matchingOffsets([{ range: a[0], criteria: a[1] }], ctx)
    return isError(offsets) ? offsets : offsets.length
  },
})
define("COUNTIFS", {
  min: 2,
  max: 254,
  errors: true,
  call: (a, ctx) => {
    if (a.length % 2) return VALUE
    const pairs = []
    for (let i = 0; i < a.length; i += 2) pairs.push({ range: a[i], criteria: scalarOf(a[i + 1], ctx) })
    const offsets = matchingOffsets(pairs, ctx)
    return isError(offsets) ? offsets : offsets.length
  },
})

function valuesAtOffsets(area: Value, shapeOf: Value, offsets: number[], ctx: Ctx): number[] | CellError {
  if (!isRef(area) || !isRef(shapeOf)) return VALUE
  const s = area.areas[0]
  const shape = shapeOf.areas[0]
  const cols = shape.c1 - shape.c0 + 1
  const out: number[] = []
  for (const o of offsets) {
    const v = ctx.cell(s.sheet, s.r0 + Math.floor(o / cols), s.c0 + (o % cols))
    if (isError(v)) return v
    if (typeof v === "number") out.push(v)
  }
  return out
}

define("AVERAGEIF", {
  min: 2,
  max: 3,
  params: ["any", "scalar", "any"],
  errors: true,
  call: (a, ctx) => {
    const offsets = matchingOffsets([{ range: a[0], criteria: a[1] }], ctx)
    if (isError(offsets)) return offsets
    const ns = valuesAtOffsets(a[2] ?? a[0], a[0], offsets, ctx)
    if (isError(ns)) return ns
    return ns.length ? excelSum(ns) / ns.length : DIV0
  },
})
function ifsPairs(a: Value[], from: number, ctx: Ctx) {
  const pairs = []
  for (let i = from; i < a.length; i += 2) pairs.push({ range: a[i], criteria: scalarOf(a[i + 1], ctx) })
  return pairs
}
function sameShape(x: Value, y: Value) {
  if (!isRef(x) || !isRef(y)) return false
  const a = x.areas[0]
  const b = y.areas[0]
  return a.r1 - a.r0 === b.r1 - b.r0 && a.c1 - a.c0 === b.c1 - b.c0
}
for (const [name, f] of [
  ["AVERAGEIFS", (ns: number[]) => (ns.length ? excelSum(ns) / ns.length : DIV0)],
  ["MAXIFS", (ns: number[]) => (ns.length ? Math.max(...ns) : 0)],
  ["MINIFS", (ns: number[]) => (ns.length ? Math.min(...ns) : 0)],
] as const) {
  define(name, {
    min: 3,
    max: 255,
    errors: true,
    call: (a, ctx) => {
      if (a.length % 2 === 0) return VALUE
      if (!sameShape(a[0], a[1])) return VALUE
      const offsets = matchingOffsets(ifsPairs(a, 1, ctx), ctx)
      if (isError(offsets)) return offsets
      const ns = valuesAtOffsets(a[0], a[1], offsets, ctx)
      return isError(ns) ? ns : f(ns)
    },
  })
}

// ── Order statistics ────────────────────────────────────────────────────────────────────────────────────────────

function nth(name: string, pick: (sorted: number[], k: number) => Scalar) {
  define(name, {
    min: 2,
    max: 2,
    params: ["any", "scalar"],
    call: (a, ctx) => {
      const ns = collectNumbers([a[0]], ctx)
      if (isError(ns)) return ns
      const k = num(a[1])
      if (isError(k)) return k
      const kk = Math.ceil(k)
      if (kk < 1 || kk > ns.length) return NUM
      return pick(
        [...ns].sort((x, y) => x - y),
        kk,
      )
    },
  })
}
nth("LARGE", (s, k) => s[s.length - k])
nth("SMALL", (s, k) => s[k - 1])

function rank(name: string, average: boolean) {
  define(name, {
    min: 2,
    max: 3,
    params: ["scalar", "any", "scalar"],
    call: (a, ctx) => {
      const x = num(a[0])
      if (isError(x)) return x
      const ns = collectNumbers([a[1]], ctx)
      if (isError(ns)) return ns
      const ascending = a.length > 2 && a[2] !== null && a[2] !== 0 && a[2] !== false
      let before = 0
      let equal = 0
      for (const n of ns) {
        if (compareScalars(n, x) === 0) equal++
        else if (ascending ? n < x : n > x) before++
      }
      if (equal === 0) return NA
      return average ? before + (equal + 1) / 2 : before + 1
    },
  })
}
rank("RANK", false)
rank("RANK.EQ", false)
rank("RANK.AVG", true)

function percentileInc(sorted: number[], p: number): Scalar {
  const n = sorted.length
  if (!n || p < 0 || p > 1) return NUM
  const h = (n - 1) * p
  const lo = Math.floor(h)
  return lo + 1 < n ? sorted[lo] + (h - lo) * (sorted[lo + 1] - sorted[lo]) : sorted[lo]
}
function percentileExc(sorted: number[], p: number): Scalar {
  const n = sorted.length
  const h = (n + 1) * p - 1
  if (!n || p <= 0 || p >= 1 || h < 0 || h > n - 1) return NUM
  const lo = Math.floor(h)
  return lo + 1 < n ? sorted[lo] + (h - lo) * (sorted[lo + 1] - sorted[lo]) : sorted[lo]
}
function sortedArg(v: Value, ctx: Ctx): number[] | CellError {
  const ns = collectNumbers([v], ctx)
  return isError(ns) ? ns : [...ns].sort((x, y) => x - y)
}
for (const [name, f] of [
  ["PERCENTILE", percentileInc],
  ["PERCENTILE.INC", percentileInc],
  ["PERCENTILE.EXC", percentileExc],
] as const) {
  define(name, {
    min: 2,
    max: 2,
    params: ["any", "scalar"],
    call: (a, ctx) => {
      const s = sortedArg(a[0], ctx)
      const p = num(a[1])
      if (isError(s)) return s
      if (isError(p)) return p
      return f(s, p)
    },
  })
}
for (const [name, f] of [
  ["QUARTILE", percentileInc],
  ["QUARTILE.INC", percentileInc],
  ["QUARTILE.EXC", percentileExc],
] as const) {
  define(name, {
    min: 2,
    max: 2,
    params: ["any", "scalar"],
    call: (a, ctx) => {
      const s = sortedArg(a[0], ctx)
      const q = int(a[1])
      if (isError(s)) return s
      if (isError(q)) return q
      if (q < 0 || q > 4) return NUM
      if (name === "QUARTILE.EXC" && (q <= 0 || q >= 4)) return NUM
      return f(s, q / 4)
    },
  })
}

/** Paired numbers from two arrays: positions where both are numbers. */
function pairs(a: Value, b: Value, ctx: Ctx): [number[], number[]] | CellError {
  const x = arrayArg(a, ctx)
  const y = arrayArg(b, ctx)
  if (isError(x)) return x
  if (isError(y)) return y
  if (x.data.length !== y.data.length) return NA
  const xs: number[] = []
  const ys: number[] = []
  for (let i = 0; i < x.data.length; i++) {
    const p = x.data[i]
    const q = y.data[i]
    if (isError(p)) return p
    if (isError(q)) return q
    if (typeof p === "number" && typeof q === "number") {
      xs.push(p)
      ys.push(q)
    }
  }
  return [xs, ys]
}
function moments(xs: number[], ys: number[]) {
  const n = xs.length
  const mx = xs.reduce((s, v) => s + v, 0) / n
  const my = ys.reduce((s, v) => s + v, 0) / n
  let sxy = 0
  let sxx = 0
  let syy = 0
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my)
    sxx += (xs[i] - mx) ** 2
    syy += (ys[i] - my) ** 2
  }
  return { n, mx, my, sxy, sxx, syy }
}
const twoArrays = (name: string, f: (m: ReturnType<typeof moments>) => Scalar, minN = 1) =>
  define(name, {
    min: 2,
    max: 2,
    call: (a, ctx) => {
      // FORECAST-style functions take (known_y, known_x).
      const p = pairs(a[0], a[1], ctx)
      if (isError(p)) return p
      if (p[0].length < minN) return DIV0
      return f(moments(p[0], p[1]))
    },
  })
twoArrays("CORREL", (m) => (m.sxx === 0 || m.syy === 0 ? DIV0 : m.sxy / Math.sqrt(m.sxx * m.syy)))
twoArrays("PEARSON", (m) => (m.sxx === 0 || m.syy === 0 ? DIV0 : m.sxy / Math.sqrt(m.sxx * m.syy)))
twoArrays("RSQ", (m) => (m.sxx === 0 || m.syy === 0 ? DIV0 : m.sxy ** 2 / (m.sxx * m.syy)))
twoArrays("COVARIANCE.P", (m) => m.sxy / m.n)
twoArrays("COVAR", (m) => m.sxy / m.n)
twoArrays("COVARIANCE.S", (m) => (m.n < 2 ? DIV0 : m.sxy / (m.n - 1)))
// SLOPE(known_y, known_x): moments() got (y, x), so x and y are swapped here.
twoArrays("SLOPE", (m) => (m.syy === 0 ? DIV0 : m.sxy / m.syy))
twoArrays("INTERCEPT", (m) => (m.syy === 0 ? DIV0 : m.mx - (m.sxy / m.syy) * m.my))
twoArrays("STEYX", (m) => (m.n < 3 || m.syy === 0 ? DIV0 : Math.sqrt((m.sxx - m.sxy ** 2 / m.syy) / (m.n - 2))))
for (const name of ["FORECAST", "FORECAST.LINEAR"]) {
  define(name, {
    min: 3,
    max: 3,
    params: ["scalar", "any", "any"],
    call: (a, ctx) => {
      const x = num(a[0])
      if (isError(x)) return x
      const p = pairs(a[1], a[2], ctx)
      if (isError(p)) return p
      const m = moments(p[0], p[1])
      if (m.n === 0 || m.syy === 0) return DIV0
      const slope = m.sxy / m.syy
      return m.mx - slope * m.my + slope * x
    },
  })
}

export { compileCriteria }
