/**
 * Depreciation (DB, DDB, VDB, AMORDEGRC, AMORLINC) and the cash-flow functions Excel added with the Analysis
 * ToolPak (CUMIPMT, CUMPRINC, DOLLARDE, DOLLARFR, FVSCHEDULE, ISPMT, MIRR, PDURATION, RRI).
 */
import { isError, type CellError } from "../../values"
import { DIV0, NUM, VALUE } from "../coerce"
import type { Value } from "../types"
import { yearFrac, dateArg } from "./datetime"
import { atpNumber } from "./engineering"
import { forEachScalar, num } from "./helpers"
import { define, scalarFn } from "./registry"

function numbers(a: Value[], count: number, defaults: number[] = []): number[] | CellError {
  const out: number[] = []
  for (let i = 0; i < count; i++) {
    const given = i < a.length && a[i] !== null
    if (!given && i - (count - defaults.length) >= 0) {
      out.push(defaults[i - (count - defaults.length)])
      continue
    }
    const v = num(a[i] ?? null)
    if (isError(v)) return v
    out.push(v)
  }
  return out
}

// ── DB, DDB, VDB ─────────────────────────────────────────────────────────────────────────────────────────────────

scalarFn("DB", 4, 5, (a) => {
  const v = numbers(a, 5, [12])
  if (isError(v)) return v
  const [cost, salvage, life, periodArg, month] = v
  const period = Math.trunc(periodArg)
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12) return NUM
  if (period > life + (month < 12 ? 1 : 0)) return NUM
  if (cost === 0) return 0
  // The rate, rounded to three decimals as Excel does.
  const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000
  let dep = (cost * rate * month) / 12
  if (period === 1) return dep
  let book = cost - dep
  for (let p = 2; p <= period; p++) {
    dep = p > life ? (book * rate * (12 - month)) / 12 : book * rate
    book -= dep
  }
  return dep
})

/** Double-declining depreciation for a (possibly fractional) period: the rate on the book value, down to salvage. */
function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life
  let before: number
  if (rate >= 1) {
    rate = 1
    before = period === 1 ? cost : 0
  } else before = cost * Math.pow(1 - rate, period - 1)
  let d = before * rate
  if (before - d < salvage) d = before - salvage
  return d < 0 ? 0 : d
}

scalarFn("DDB", 4, 5, (a) => {
  const v = numbers(a, 5, [2])
  if (isError(v)) return v
  const [cost, salvage, life, period, factor] = v
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0) return NUM
  return ddb(cost, salvage, life, period, factor)
})

/**
 * VDB with the switch to straight line: Excel splits time at the whole periods before `start`, then at `start`, `start
 * + 1`, … up to `end`; each piece takes its length × rate × the book value (or its share of straight line over the
 * life left, when that is more), never going below salvage.
 */
function vdbSwitch(cost: number, salvage: number, life: number, start: number, end: number, factor: number) {
  const rate = factor / life
  const points: number[] = []
  for (let i = 0; i < start; i++) points.push(i)
  points.push(start)
  for (let x = start + 1; x < end; x++) points.push(x)
  if (end > start) points.push(end)
  let book = cost
  let total = 0
  for (let i = 0; i + 1 < points.length; i++) {
    const from = points[i]
    const t = points[i + 1] - from
    let d = t * rate * book
    const straight = (t * (book - salvage)) / (life - from)
    if (straight > d) d = straight
    if (book - d < salvage) d = book - salvage
    if (from >= start) total += d
    book -= d
  }
  return total
}

scalarFn("VDB", 5, 7, (a) => {
  const v = numbers(a, 6, [2])
  if (isError(v)) return v
  const [cost, salvage, life, start, end, factor] = v
  let noSwitch = false
  if (a.length > 6 && a[6] !== null) {
    const s = num(a[6])
    if (isError(s)) return s
    noSwitch = s !== 0
  }
  if (start < 0 || end < start || end > life || cost < 0 || salvage < 0 || factor < 0 || life <= 0) return NUM
  if (!noSwitch) return vdbSwitch(cost, salvage, life, start, end, factor)
  // Without the switch: whole-period DDB amounts, the first and last weighted by the part of them inside.
  const intStart = Math.floor(start)
  const intEnd = Math.ceil(end)
  let total = 0
  for (let i = intStart + 1; i <= intEnd; i++) {
    let term = ddb(cost, salvage, life, i, factor)
    if (i === intStart + 1) term *= Math.min(end, intStart + 1) - start
    else if (i === intEnd) term *= end + 1 - intEnd
    total += term
  }
  return total
})

// ── French depreciation (AMORDEGRC, AMORLINC) ────────────────────────────────────────────────────────────────────

function amorArgs(a: Value[]): number[] | CellError {
  const out: number[] = []
  for (let i = 0; i < 6; i++) {
    const v = i === 1 || i === 2 ? (typeof a[i] === "boolean" ? VALUE : dateArg(a[i])) : atpNumber(a[i])
    if (isError(v)) return v
    out.push(i === 1 || i === 2 ? Math.floor(v) : v)
  }
  const b = a.length > 6 && a[6] !== null ? atpNumber(a[6]) : 0
  if (isError(b)) return b
  const basis = Math.trunc(b)
  // Actual/360 isn't allowed for these.
  if (basis < 0 || basis > 4 || basis === 2) return NUM
  out.push(basis)
  return out
}
const roundHalfUp = (x: number) => Math.sign(x) * Math.round(Math.abs(x))

scalarFn("AMORDEGRC", 6, 7, (a) => {
  const v = amorArgs(a)
  if (isError(v)) return v
  const [, bought, firstPeriod, salvage, periodArg, , basis] = v
  let [cost, , , , , rate] = v
  if (bought > firstPeriod || cost < 0 || salvage < 0 || salvage > cost || periodArg < 0 || rate <= 0) return NUM
  const period = Math.trunc(periodArg)
  // The coefficient goes by the life rounded up to whole years; a life of up to two years is #NUM!.
  const years = Math.ceil(1 / rate)
  if (years <= 2) return NUM
  rate *= years <= 4 ? 1.5 : years <= 6 ? 2 : 2.5
  let dep = roundHalfUp(yearFrac(bought, firstPeriod, basis) * rate * cost)
  cost -= dep
  let rest = cost - salvage
  if (period === 0) return dep
  for (let n = 0; n < period; n++) {
    dep = roundHalfUp(rate * cost)
    rest -= dep
    // Half of what is left in the period two before the last (or once the rate would pass salvage), then nothing.
    if (rest < 0 || n + 1 === years - 2) return period - n <= 1 ? roundHalfUp(cost * 0.5) : 0
    cost -= dep
  }
  return dep
})

scalarFn("AMORLINC", 6, 7, (a) => {
  const v = amorArgs(a)
  if (isError(v)) return v
  const [cost, bought, firstPeriod, salvage, periodArg, rate, basis] = v
  if (bought > firstPeriod || cost < 0 || salvage < 0 || salvage > cost || periodArg < 0 || rate <= 0) return NUM
  const period = Math.trunc(periodArg)
  const full = cost * rate
  const first = yearFrac(bought, firstPeriod, basis) * rate * cost
  const wholePeriods = Math.trunc((cost - salvage - first) / full)
  if (period === 0) return first
  if (period <= wholePeriods) return full
  if (period === wholePeriods + 1) return cost - salvage - full * wholePeriods - first
  return 0
})

// ── Cumulative interest and principal ────────────────────────────────────────────────────────────────────────────

function pmt(rate: number, nper: number, pv: number, type: number): number {
  const f = Math.pow(1 + rate, nper)
  return (-pv * f * rate) / ((1 + rate * type) * (f - 1))
}
function fv(rate: number, nper: number, payment: number, pv: number, type: number): number {
  const f = Math.pow(1 + rate, nper)
  return -(pv * f + (payment * (1 + rate * type) * (f - 1)) / rate)
}

for (const principal of [false, true]) {
  scalarFn(principal ? "CUMPRINC" : "CUMIPMT", 6, 6, (a) => {
    const v = numbers(a, 6)
    if (isError(v)) return v
    const [rate, nper, pv, startArg, endArg, type] = v
    const start = Math.ceil(startArg)
    const end = Math.trunc(endArg)
    if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper || (type !== 0 && type !== 1))
      return NUM
    const payment = pmt(rate, nper, pv, type)
    let interest = 0
    let paid = 0
    let first = start
    if (start === 1) {
      if (type === 0) {
        interest = -pv
        paid = payment + pv * rate
      } else paid = payment
      first = 2
    }
    for (let i = first; i <= end; i++) {
      const balance = type === 1 ? fv(rate, i - 2, payment, pv, 1) - payment : fv(rate, i - 1, payment, pv, 0)
      interest += balance
      paid += payment - balance * rate
    }
    return principal ? paid : interest * rate
  })
}

// ── DOLLARDE, DOLLARFR ───────────────────────────────────────────────────────────────────────────────────────────

for (const toDecimal of [true, false]) {
  scalarFn(toDecimal ? "DOLLARDE" : "DOLLARFR", 2, 2, (a) => {
    const x = atpNumber(a[0])
    if (isError(x)) return x
    const f = atpNumber(a[1])
    if (isError(f)) return f
    const fraction = Math.trunc(f)
    if (fraction < 0) return NUM
    if (fraction === 0) return DIV0
    const whole = Math.trunc(x)
    const scale = Math.pow(10, Math.ceil(Math.log10(fraction)))
    const rest = x - whole
    return whole + (toDecimal ? (rest * scale) / fraction : (rest * fraction) / scale)
  })
}

// ── FVSCHEDULE, ISPMT, MIRR, PDURATION, RRI ──────────────────────────────────────────────────────────────────────

define("FVSCHEDULE", {
  min: 2,
  max: 2,
  call: (a, ctx) => {
    const principal = atpNumber(a[0] as Value)
    if (isError(principal)) return principal
    let value = principal
    let error: CellError | null = null
    forEachScalar([a[1]], ctx, (r) => {
      if (isError(r)) error = r
      else if (typeof r !== "number") error = VALUE
      else value *= 1 + r
      return error === null
    })
    return error ?? value
  },
})

scalarFn("ISPMT", 4, 4, (a) => {
  const v = numbers(a, 4)
  if (isError(v)) return v
  const [rate, per, nper, pv] = v
  if (nper === 0) return DIV0
  return pv * rate * (per / nper - 1)
})

define("MIRR", {
  min: 3,
  max: 3,
  params: ["any", "scalar", "scalar"],
  call: (a, ctx) => {
    const values: number[] = []
    let error: CellError | null = null
    forEachScalar([a[0]], ctx, (v) => {
      if (isError(v)) error = v
      else if (typeof v === "number") values.push(v)
      return error === null
    })
    if (error) return error
    const finance = num(a[1])
    if (isError(finance)) return finance
    const reinvest = num(a[2])
    if (isError(reinvest)) return reinvest
    const n = values.length
    let positive = 0
    let negative = 0
    values.forEach((v, i) => {
      if (v > 0) positive += v / Math.pow(1 + reinvest, i)
      else if (v < 0) negative += v / Math.pow(1 + finance, i)
    })
    if (negative === 0 || n < 2) return DIV0
    return Math.pow((-positive * Math.pow(1 + reinvest, n - 1)) / negative, 1 / (n - 1)) - 1
  },
})

scalarFn("PDURATION", 3, 3, (a) => {
  const v = numbers(a, 3)
  if (isError(v)) return v
  const [rate, pv, fvalue] = v
  if (rate <= 0 || pv <= 0 || fvalue <= 0) return NUM
  return (Math.log(fvalue) - Math.log(pv)) / Math.log(1 + rate)
})

scalarFn("RRI", 3, 3, (a) => {
  const v = numbers(a, 3)
  if (isError(v)) return v
  const [nper, pv, fvalue] = v
  if (nper <= 0 || pv === 0) return NUM
  const r = Math.pow(fvalue / pv, 1 / nper) - 1
  return Number.isFinite(r) ? r : NUM
})
