import { isError, type CellError } from "../../values"
import { DIV0, NUM, VALUE } from "../coerce"
import { squarePower, toMatrix } from "../evaluate"
import type { Value } from "../types"
import { collectNumbers, num } from "./helpers"
import { define, scalarFn } from "./registry"

/** Optional numeric argument with a default. */
function opt(a: Value[], i: number, fallback: number): number | CellError {
  return a.length > i && a[i] !== null ? num(a[i]) : fallback
}

function nums(a: Value[], defaults: number[]): number[] | CellError {
  const out: number[] = []
  for (let i = 0; i < defaults.length; i++) {
    const v = Number.isNaN(defaults[i]) ? num(a[i]) : opt(a, i, defaults[i])
    if (isError(v)) return v
    out.push(v)
  }
  return out
}

/**
 * (1 + rate)^n. Checked against Excel to the last bit: PMT, IPMT and PPMT compute it as exp(n·log1p(rate)); FV and
 * PV as a power (by squaring for whole n, like the ^ operator).
 */
function growth(rate: number, n: number) {
  return Math.exp(n * Math.log1p(rate))
}
function compound(rate: number, n: number) {
  return Number.isInteger(n) ? squarePower(1 + rate, n) : (1 + rate) ** n
}

/** The time-value-of-money identity that PMT, PV, FV, NPER and RATE solve. */
function fvOf(rate: number, nper: number, pmt: number, pv: number, type: number, grow = growth) {
  if (rate === 0) return -(pv + pmt * nper)
  const f = grow(rate, nper)
  return -(pv * f + (pmt * (1 + rate * type) * (f - 1)) / rate)
}
function pmtOf(rate: number, nper: number, pv: number, fv: number, type: number) {
  if (rate === 0) return -(pv + fv) / nper
  const f = growth(rate, nper)
  return (-(pv * f + fv) * rate) / ((1 + rate * type) * (f - 1))
}

const R = NaN // required

scalarFn("PMT", 3, 5, (a) => {
  const v = nums(a, [R, R, R, 0, 0])
  if (isError(v)) return v
  const [rate, nper, pv, fv, type] = v
  if (nper === 0) return NUM
  return pmtOf(rate, nper, pv, fv, type ? 1 : 0)
})
scalarFn("FV", 3, 5, (a) => {
  const v = nums(a, [R, R, R, 0, 0])
  if (isError(v)) return v
  const [rate, nper, pmt, pv, type] = v
  return fvOf(rate, nper, pmt, pv, type ? 1 : 0, compound)
})
scalarFn("PV", 3, 5, (a) => {
  const v = nums(a, [R, R, R, 0, 0])
  if (isError(v)) return v
  const [rate, nper, pmt, fv, type] = v
  const t = type ? 1 : 0
  if (rate === 0) return -(fv + pmt * nper)
  const f = compound(rate, nper)
  return -(fv + (pmt * (1 + rate * t) * (f - 1)) / rate) / f
})
scalarFn("NPER", 3, 5, (a) => {
  const v = nums(a, [R, R, R, 0, 0])
  if (isError(v)) return v
  const [rate, pmt, pv, fv, type] = v
  const t = type ? 1 : 0
  if (rate === 0) {
    if (pmt === 0) return DIV0
    return -(pv + fv) / pmt
  }
  const num1 = pmt * (1 + rate * t) - fv * rate
  const den = pv * rate + pmt * (1 + rate * t)
  if (num1 / den <= 0) return NUM
  return Math.log(num1 / den) / Math.log(1 + rate)
})
scalarFn("IPMT", 4, 6, (a) => {
  const v = nums(a, [R, R, R, R, 0, 0])
  if (isError(v)) return v
  const [rate, per, nper, pv, fv, type] = v
  if (per < 1 || per > nper) return NUM
  const t = type ? 1 : 0
  const pmt = pmtOf(rate, nper, pv, fv, t)
  if (t === 1 && per === 1) return 0
  const fvBefore = fvOf(rate, per - 1, pmt, pv, t)
  return t === 1 ? (fvBefore - pmt) * rate : fvBefore * rate
})
scalarFn("PPMT", 4, 6, (a) => {
  const v = nums(a, [R, R, R, R, 0, 0])
  if (isError(v)) return v
  const [rate, per, nper, pv, fv, type] = v
  if (per < 1 || per > nper) return NUM
  const t = type ? 1 : 0
  const pmt = pmtOf(rate, nper, pv, fv, t)
  let interest = 0
  if (!(t === 1 && per === 1)) {
    const fvBefore = fvOf(rate, per - 1, pmt, pv, t)
    interest = t === 1 ? (fvBefore - pmt) * rate : fvBefore * rate
  }
  return pmt - interest
})

scalarFn("RATE", 3, 6, (a) => {
  const v = nums(a, [R, R, R, 0, 0, 0.1])
  if (isError(v)) return v
  const [nper, pmt, pv, fv, type, guess] = v
  const t = type ? 1 : 0
  // Newton's method on f(rate) = FV identity, as Excel does (20 iterations, 1e-7 tolerance).
  let rate = guess
  for (let i = 0; i < 50; i++) {
    const f = rateEquation(rate, nper, pmt, pv, fv, t)
    const h = 1e-7
    const d = (rateEquation(rate + h, nper, pmt, pv, fv, t) - f) / h
    if (d === 0 || !Number.isFinite(d)) return NUM
    const next = rate - f / d
    if (!Number.isFinite(next)) return NUM
    if (Math.abs(next - rate) < 1e-10) return next
    rate = next
  }
  return NUM
})
function rateEquation(rate: number, nper: number, pmt: number, pv: number, fv: number, type: number) {
  if (Math.abs(rate) < 1e-12) return pv + pmt * nper + fv
  const f = growth(rate, nper)
  return pv * f + (pmt * (1 + rate * type) * (f - 1)) / rate + fv
}

define("NPV", {
  min: 2,
  max: 255,
  call: (a, ctx) => {
    const rate = num(a[0])
    if (isError(rate)) return rate
    const ns = collectNumbers(a.slice(1), ctx)
    if (isError(ns)) return ns
    if (rate === -1) return DIV0
    return ns.reduce((s, x, i) => s + x / (1 + rate) ** (i + 1), 0)
  },
})

function irrOf(values: number[], guess: number, dates?: number[]): number | CellError {
  const t = (i: number) => (dates ? (dates[i] - dates[0]) / 365 : i)
  const f = (r: number) => values.reduce((s, v, i) => s + v / (1 + r) ** t(i), 0)
  const df = (r: number) => values.reduce((s, v, i) => s - (t(i) * v) / (1 + r) ** (t(i) + 1), 0)
  if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) return NUM
  let r = guess
  for (let i = 0; i < 100; i++) {
    const d = df(r)
    if (d === 0) return NUM
    const next = r - f(r) / d
    if (!Number.isFinite(next) || next <= -1) return NUM
    if (Math.abs(next - r) < 1e-10) return next
    r = next
  }
  return NUM
}
define("IRR", {
  min: 1,
  max: 2,
  call: (a, ctx) => {
    const m = toMatrix(a[0], ctx)
    if (isError(m)) return m
    const values = m.data.filter((v): v is number => typeof v === "number")
    const guess = a.length > 1 && a[1] !== null ? num(a[1]) : 0.1
    if (isError(guess)) return guess
    return irrOf(values, guess)
  },
})
function pairedCashflows(a: Value[], ctx: Parameters<typeof toMatrix>[1]) {
  const v = toMatrix(a[0], ctx)
  const d = toMatrix(a[1], ctx)
  if (isError(v)) return v
  if (isError(d)) return d
  if (v.data.length !== d.data.length) return NUM
  const values: number[] = []
  const dates: number[] = []
  for (let i = 0; i < v.data.length; i++) {
    const x = v.data[i]
    const y = d.data[i]
    if (isError(x)) return x
    if (isError(y)) return y
    if (typeof x !== "number" || typeof y !== "number") return VALUE
    values.push(x)
    dates.push(Math.floor(y))
  }
  if (dates.some((x) => x < dates[0])) return NUM
  return { values, dates }
}
define("XNPV", {
  min: 3,
  max: 3,
  call: (a, ctx) => {
    const rate = num(a[0])
    if (isError(rate)) return rate
    const p = pairedCashflows(a.slice(1), ctx)
    if (isError(p)) return p
    return p.values.reduce((s, x, i) => s + x / (1 + rate) ** ((p.dates[i] - p.dates[0]) / 365), 0)
  },
})
define("XIRR", {
  min: 2,
  max: 3,
  call: (a, ctx) => {
    const p = pairedCashflows(a, ctx)
    if (isError(p)) return p
    const guess = a.length > 2 && a[2] !== null ? num(a[2]) : 0.1
    if (isError(guess)) return guess
    return irrOf(p.values, guess, p.dates)
  },
})
scalarFn("SLN", 3, 3, (a) => {
  const v = nums(a, [R, R, R])
  if (isError(v)) return v
  const [cost, salvage, life] = v
  return life === 0 ? DIV0 : (cost - salvage) / life
})
scalarFn("SYD", 4, 4, (a) => {
  const v = nums(a, [R, R, R, R])
  if (isError(v)) return v
  const [cost, salvage, life, per] = v
  if (life <= 0 || per <= 0 || per > life) return NUM
  return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1))
})
scalarFn("EFFECT", 2, 2, (a) => {
  const v = nums(a, [R, R])
  if (isError(v)) return v
  const [rate, n] = v
  const periods = Math.trunc(n)
  if (rate <= 0 || periods < 1) return NUM
  return squarePower(1 + rate / periods, periods) - 1
})
scalarFn("NOMINAL", 2, 2, (a) => {
  const v = nums(a, [R, R])
  if (isError(v)) return v
  const [rate, n] = v
  const periods = Math.trunc(n)
  if (rate <= 0 || periods < 1) return NUM
  return periods * ((1 + rate) ** (1 / periods) - 1)
})
export { VALUE }
