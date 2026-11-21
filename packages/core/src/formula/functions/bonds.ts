/**
 * Securities: coupon dates (COUP*), coupon bonds (PRICE, YIELD, DURATION, MDURATION, ACCRINT), discounted and
 * at-maturity securities (DISC, INTRATE, RECEIVED, PRICEDISC, YIELDDISC, PRICEMAT, YIELDMAT, ACCRINTM), Treasury
 * bills and bonds with an odd first or last period (ODDF*, ODDL*).
 *
 * Coupon dates run back from maturity by 12/frequency months, keeping to the month's end when maturity is a month's
 * last day. Days are counted by the basis: 0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European
 * 30/360. The date arithmetic follows the Analysis ToolPak (as also reproduced by LibreOffice's ScaDate).
 */
import { isError, type CellError } from "../../values"
import { NUM, VALUE } from "../coerce"
import { dateSerial, daysInMonth, isLeapYear, serialToDate } from "../dates"
import type { Value } from "../types"
import { dateArg, days360, yearFrac } from "./datetime"
import { atpNumber } from "./engineering"
import { scalarFn } from "./registry"

// ── Arguments ────────────────────────────────────────────────────────────────────────────────────────────────────

/** A date argument: a whole serial (times are dropped); a logical is #VALUE!. */
function date(v: Value): number | CellError {
  if (typeof v === "boolean") return VALUE
  const d = dateArg(v)
  return isError(d) ? d : Math.floor(d)
}
function number(v: Value): number | CellError {
  return atpNumber(v)
}
function optional(a: Value[], i: number, fallback: number): number | CellError {
  return a.length > i && a[i] !== null ? number(a[i]) : fallback
}
function basisArg(a: Value[], i: number): number | CellError {
  const b = optional(a, i, 0)
  if (isError(b)) return b
  const basis = Math.trunc(b)
  return basis < 0 || basis > 4 ? NUM : basis
}
function frequencyArg(v: Value): number | CellError {
  const f = number(v)
  if (isError(f)) return f
  const freq = Math.trunc(f)
  return freq === 1 || freq === 2 || freq === 4 ? freq : NUM
}

/** Reads the arguments by kind ("d" date, "n" number, "f" frequency, "b" optional basis); the first error wins. */
function read(a: Value[], kinds: string): number[] | CellError {
  const out: number[] = []
  for (let i = 0; i < kinds.length; i++) {
    const k = kinds[i]
    const v = k === "d" ? date(a[i]) : k === "f" ? frequencyArg(a[i]) : k === "b" ? basisArg(a, i) : number(a[i])
    if (isError(v)) return v
    out.push(v)
  }
  return out
}

// ── Dates as the coupon functions see them ───────────────────────────────────────────────────────────────────────

/** A date that remembers its original day (and whether it was a month's last), counted in 30-day months or not. */
class CouponDate {
  origDay = 1
  day = 1
  month = 1
  year = 1900
  lastDay = false
  days30 = false
  us = false

  static of(serial: number, basis: number): CouponDate {
    const d = new CouponDate()
    const p = serialToDate(serial)
    d.origDay = p.day
    d.month = p.month
    d.year = p.year
    d.lastDay = p.day >= daysInMonth(p.year, p.month)
    d.days30 = basis === 0 || basis === 4
    d.us = basis === 0
    d.setDay()
    return d
  }
  clone(): CouponDate {
    return Object.assign(new CouponDate(), this)
  }
  private setDay() {
    if (this.days30) {
      this.day = Math.min(this.origDay, 30)
      if (this.lastDay || this.day >= daysInMonth(this.year, this.month)) this.day = 30
    } else {
      const last = daysInMonth(this.year, this.month)
      this.day = this.lastDay ? last : Math.min(this.origDay, last)
    }
  }
  addMonths(n: number) {
    let m = this.month + n
    if (m > 12) {
      m--
      this.year += Math.floor(m / 12)
      this.month = (m % 12) + 1
    } else if (m < 1) {
      this.year += Math.trunc(m / 12) - 1
      this.month = (m % 12) + 12
    } else this.month = m
    this.setDay()
  }
  addYears(n: number) {
    this.year += n
    this.setDay()
  }
  setYear(y: number) {
    this.year = y
    this.setDay()
  }
  serial(): number {
    const last = daysInMonth(this.year, this.month)
    return dateSerial(this.year, this.month, this.lastDay ? last : Math.min(last, this.origDay))
  }
  compare(o: CouponDate): number {
    if (this.year !== o.year) return this.year - o.year
    if (this.month !== o.month) return this.month - o.month
    if (this.day !== o.day) return this.day - o.day
    if (this.lastDay || o.lastDay) return this.lastDay === o.lastDay ? 0 : this.lastDay ? 1 : -1
    return this.origDay - o.origDay
  }
  private monthDays(month = this.month) {
    return this.days30 ? 30 : daysInMonth(this.year, month)
  }
  private monthRange(from: number, to: number) {
    if (from > to) return 0
    if (this.days30) return (to - from + 1) * 30
    let n = 0
    for (let m = from; m <= to; m++) n += this.monthDays(m)
    return n
  }
  private yearRange(from: number, to: number) {
    if (from > to) return 0
    if (this.days30) return (to - from + 1) * 360
    let n = 0
    for (let y = from; y <= to; y++) n += isLeapYear(y) ? 366 : 365
    return n
  }
  /** Days from one date to another in the basis's count. */
  static diff(fromDate: CouponDate, toDate: CouponDate): number {
    if (fromDate.compare(toDate) > 0) return CouponDate.diff(toDate, fromDate)
    const from = fromDate.clone()
    const to = toDate.clone()
    if (toDate.days30) {
      if (toDate.us) {
        if ((fromDate.month === 2 || fromDate.day < 30) && to.origDay === 31) to.day = 31
        else if (to.month === 2 && to.lastDay) to.day = daysInMonth(to.year, 2)
      } else {
        if (from.month === 2 && from.day === 30) from.day = daysInMonth(from.year, 2)
        if (to.month === 2 && to.day === 30) to.day = daysInMonth(to.year, 2)
      }
    }
    let n = 0
    if (from.year < to.year || (from.year === to.year && from.month < to.month)) {
      n = from.monthDays() - from.day + 1
      from.origDay = from.day = 1
      from.lastDay = false
      from.addMonths(1)
      if (from.year < to.year) {
        n += from.monthRange(from.month, 12)
        from.addMonths(13 - from.month)
        n += from.yearRange(from.year, to.year - 1)
        from.addYears(to.year - from.year)
      }
      n += from.monthRange(from.month, to.month - 1)
      from.addMonths(to.month - from.month)
    }
    n += to.day - from.day
    return n > 0 ? n : 0
  }
}

type Coupons = { settle: CouponDate; mat: CouponDate; freq: number; basis: number }

function coupons(settle: number, mat: number, freq: number, basis: number): Coupons | CellError {
  if (settle >= mat) return NUM
  return { settle: CouponDate.of(settle, basis), mat: CouponDate.of(mat, basis), freq, basis }
}
/** The coupon date on or before settlement. */
function previousCoupon(c: Coupons): CouponDate {
  const d = c.mat.clone()
  d.setYear(c.settle.year)
  if (d.compare(c.settle) < 0) d.addYears(1)
  while (d.compare(c.settle) > 0) d.addMonths(-12 / c.freq)
  return d
}
/** The coupon date after settlement. */
function nextCoupon(c: Coupons): CouponDate {
  const d = c.mat.clone()
  d.setYear(c.settle.year)
  if (d.compare(c.settle) > 0) d.addYears(-1)
  while (d.compare(c.settle) <= 0) d.addMonths(12 / c.freq)
  return d
}
function coupNum(c: Coupons): number {
  const pcd = previousCoupon(c)
  const months = (c.mat.year - pcd.year) * 12 + c.mat.month - pcd.month
  return (months * c.freq) / 12
}
function coupDayBs(c: Coupons): number {
  return CouponDate.diff(previousCoupon(c), c.settle)
}
function coupDays(c: Coupons): number {
  if (c.basis === 1) {
    // Excel counts from the previous coupon to one period later on the same day (not the month's end).
    const pcd = previousCoupon(c).serial()
    return shiftMonths(pcd, 12 / c.freq, false) - pcd
  }
  return (c.basis === 3 ? 365 : 360) / c.freq
}
function coupDaysNc(c: Coupons): number {
  if (c.basis === 0) return coupDays(c) - coupDayBs(c)
  const settle = c.settle.serial()
  const ncd = nextCoupon(c).serial()
  return c.basis === 4 ? days360(parts(settle), parts(ncd), true) : ncd - settle
}

function couponFn(name: string, f: (c: Coupons) => number) {
  scalarFn(name, 3, 4, (a) => {
    const v = read(a, "ddfb")
    if (isError(v)) return v
    const c = coupons(v[0], v[1], v[2], v[3])
    return isError(c) ? c : f(c)
  })
}
couponFn("COUPDAYBS", coupDayBs)
couponFn("COUPDAYS", coupDays)
couponFn("COUPDAYSNC", coupDaysNc)
couponFn("COUPNCD", (c) => nextCoupon(c).serial())
couponFn("COUPPCD", (c) => previousCoupon(c).serial())
couponFn("COUPNUM", coupNum)

// ── Coupon bonds ─────────────────────────────────────────────────────────────────────────────────────────────────

function price(c: Coupons, rate: number, yld: number, redemption: number): number {
  const f = c.freq
  const e = coupDays(c)
  const n = coupNum(c)
  const a = coupDayBs(c)
  // Excel takes the days to the next coupon as the period's length less the days already accrued.
  const dsc = (e - a) / e
  const coupon = (100 * rate) / f
  // In the last period, simple interest to redemption.
  if (n <= 1) return (redemption + coupon) / (1 + (dsc * yld) / f) - (coupon * a) / e
  const base = 1 + yld / f
  let p = redemption / Math.pow(base, n - 1 + dsc)
  for (let k = 0; k < n; k++) p += coupon / Math.pow(base, k + dsc)
  return p - (coupon * a) / e
}

scalarFn("PRICE", 6, 7, (a) => {
  const v = read(a, "ddnnnfb")
  if (isError(v)) return v
  const [settle, mat, rate, yld, redemption, freq, basis] = v
  if (rate < 0 || yld < 0 || redemption <= 0) return NUM
  const c = coupons(settle, mat, freq, basis)
  return isError(c) ? c : price(c, rate, yld, redemption)
})

/**
 * The yield that gives a price: solved (price falls as yield rises), or in the last coupon period in closed form, with
 * the days to redemption in the basis's count (actual/360 and actual/365 count this one as actual/actual).
 */
function yieldOf(c: Coupons, rate: number, pr: number, redemption: number): number | CellError {
  if (coupNum(c) > 1) return solveDecreasing((y) => price(c, rate, y, redemption), pr, -c.freq)
  const settle = c.settle.serial()
  const mat = c.mat.serial()
  const cc = c.basis === 2 || c.basis === 3 ? (coupons(settle, mat, c.freq, 1) as Coupons) : c
  const e = coupDays(cc)
  const a = coupDayBs(cc)
  const dsr = c.basis === 0 || c.basis === 4 ? days360(parts(settle), parts(mat), c.basis === 4) : mat - settle
  const paid = pr / 100 + ((a / e) * rate) / c.freq
  return ((redemption / 100 + rate / c.freq - paid) / paid) * ((c.freq * e) / dsr)
}

/**
 * The x above `floor` where a decreasing f(x) = target: bracketed (doubling up from 0, or halving the way down to the
 * floor for negative yields), then narrowed by the secant method with bisection steps to full precision.
 */
function solveDecreasing(f: (x: number) => number, target: number, floor = 0): number | CellError {
  let lo = 0
  let hi = 1
  let flo = f(lo) - target
  if (flo === 0) return lo
  let fhi: number
  if (flo < 0) {
    // The price at 0 is already below target: the yield is negative.
    hi = 0
    fhi = flo
    for (let k = 1; ; k++) {
      if (k > 60 || floor === 0) return NUM
      lo = floor * (1 - Math.pow(2, -k))
      flo = f(lo) - target
      if (flo > 0) break
      hi = lo
      fhi = flo
    }
  } else {
    fhi = f(hi) - target
    for (let i = 0; fhi > 0; i++) {
      if (i > 60) return NUM
      lo = hi
      flo = fhi
      hi *= 2
      fhi = f(hi) - target
    }
  }
  for (let i = 0; i < 200; i++) {
    // Secant inside the bracket, bisection when it would leave it.
    let x = hi - (fhi * (hi - lo)) / (fhi - flo)
    if (!(x > lo && x < hi)) x = (lo + hi) / 2
    const fx = f(x) - target
    if (fx === 0 || hi - lo <= Math.max(Math.abs(x), 1e-300) * 1e-16) return x
    if (fx > 0) {
      lo = x
      flo = fx
    } else {
      hi = x
      fhi = fx
    }
    if (i % 3 === 2) {
      // Keep the bracket shrinking even when the secant creeps from one side.
      const mid = (lo + hi) / 2
      const fm = f(mid) - target
      if (fm > 0) {
        lo = mid
        flo = fm
      } else {
        hi = mid
        fhi = fm
      }
    }
  }
  return (lo + hi) / 2
}

scalarFn("YIELD", 6, 7, (a) => {
  const v = read(a, "ddnnnfb")
  if (isError(v)) return v
  const [settle, mat, rate, pr, redemption, freq, basis] = v
  if (rate < 0 || pr <= 0 || redemption <= 0) return NUM
  const c = coupons(settle, mat, freq, basis)
  return isError(c) ? c : yieldOf(c, rate, pr, redemption)
})

/** Macaulay duration in years: the cash flows' times (in periods, as PRICE counts them) weighted by their value. */
function duration(c: Coupons, coupon: number, yld: number): number {
  const e = coupDays(c)
  const n = coupNum(c)
  const dsc = (e - coupDayBs(c)) / e
  const cash = (coupon * 100) / c.freq
  const base = yld / c.freq + 1
  let dur = 0
  let p = 0
  for (let k = 1; k <= n; k++) {
    const t = k - 1 + dsc
    const flow = k === n ? cash + 100 : cash
    const v = flow / Math.pow(base, t)
    dur += t * v
    p += v
  }
  return dur / p / c.freq
}

for (const modified of [false, true]) {
  scalarFn(modified ? "MDURATION" : "DURATION", 5, 6, (a) => {
    const v = read(a, "ddnnfb")
    if (isError(v)) return v
    const [settle, mat, coupon, yld, freq, basis] = v
    if (coupon < 0 || yld < 0) return NUM
    const c = coupons(settle, mat, freq, basis)
    if (isError(c)) return c
    const d = duration(c, coupon, yld)
    return modified ? d / (1 + yld / freq) : d
  })
}

// ── Accrued interest ─────────────────────────────────────────────────────────────────────────────────────────────

function lastOfMonth(serial: number): boolean {
  const p = serialToDate(serial)
  return p.day === daysInMonth(p.year, p.month)
}
/** A date some months away; at the month's end when `endOfMonth`. */
function shiftMonths(serial: number, months: number, endOfMonth: boolean): number {
  const p = serialToDate(serial)
  const m = p.month - 1 + months
  const year = p.year + Math.floor(m / 12)
  const month = (((m % 12) + 12) % 12) + 1
  const last = daysInMonth(year, month)
  return dateSerial(year, month, endOfMonth ? last : Math.min(p.day, last))
}
function parts(serial: number) {
  return serialToDate(serial)
}

/** 30/360 US for accrued interest: only the start date's day is changed when `startOnly` (the issue's own period). */
function days360Us(start: number, end: number, startOnly: boolean): number {
  const s = parts(start)
  const e = parts(end)
  let sd = s.day
  let ed = e.day
  const lastFeb = (d: typeof s) => d.month === 2 && d.day === daysInMonth(d.year, 2)
  if (lastFeb(s)) {
    if (!startOnly && lastFeb(e)) ed = 30
    sd = 30
  }
  if (ed === 31 && sd >= 30) ed = 30
  if (sd === 31) sd = 30
  return (e.year - s.year) * 360 + (e.month - s.month) * 30 + (ed - sd)
}
function daysBetween(start: number, end: number, basis: number): number {
  if (basis === 0) return days360Us(start, end, false)
  if (basis === 4) return days360(parts(start), parts(end), true)
  return end - start
}
/** The length of a coupon period in the basis's days. */
function periodDays(pcd: number, ncd: number, basis: number, freq: number): number {
  if (basis === 1) return ncd - pcd
  if (basis === 3) return 365 / freq
  return 360 / freq
}

scalarFn("ACCRINT", 6, 8, (a) => {
  const v = read(a, "dddn")
  if (isError(v)) return v
  const [issue, firstInterest, settle, rate] = v
  const par = optional(a, 4, 1000)
  if (isError(par)) return par
  const freq = frequencyArg(a[5])
  if (isError(freq)) return freq
  const basis = basisArg(a, 6)
  if (isError(basis)) return basis
  let byPeriods = true
  if (a.length > 7 && a[7] !== null) {
    const m = a[7]
    if (typeof m === "string") return VALUE
    const n = number(typeof m === "boolean" ? (m ? 1 : 0) : m)
    if (isError(n)) return n
    byPeriods = n !== 0
  }
  if (rate <= 0 || par <= 0 || issue >= settle) return NUM
  const months = 12 / freq
  const eom = lastOfMonth(firstInterest)
  // The first coupon period sets the length the current period's days are divided by.
  const firstStart = shiftMonths(firstInterest, -months, eom)
  let pcd = firstStart
  if (settle > firstInterest && byPeriods) {
    pcd = firstInterest
    for (let next = firstInterest; next < settle; next = shiftMonths(next, months, eom)) pcd = next
  }
  const from = Math.max(issue, pcd)
  let total = daysBetween(from, settle, basis) / periodDays(firstStart, firstInterest, basis, freq)
  // The periods before, back to the issue: whole ones count 1 (by periods) or nothing, the issue's own one in part.
  for (let ncd = pcd; ncd > issue;) {
    const start = shiftMonths(ncd, -months, eom)
    if (issue <= start) total += byPeriods ? 1 : 0
    else {
      const days =
        basis === 0 ? days360Us(issue, ncd, true) : basis === 4 ? days360(parts(issue), parts(ncd), true) : ncd - issue
      const length =
        basis === 0
          ? days360Us(start, ncd, false)
          : basis === 4
            ? days360(parts(start), parts(ncd), true)
            : periodDays(start, ncd, basis, freq)
      total += days / length
    }
    ncd = start
  }
  return ((par * rate) / freq) * total
})

scalarFn("ACCRINTM", 4, 5, (a) => {
  const v = read(a, "ddn")
  if (isError(v)) return v
  const [issue, settle, rate] = v
  const par = optional(a, 3, 1000)
  if (isError(par)) return par
  const basis = basisArg(a, 4)
  if (isError(basis)) return basis
  if (rate <= 0 || par <= 0 || issue >= settle) return NUM
  return par * rate * yearFrac(issue, settle, basis)
})

// ── Discounted securities and securities paying at maturity ─────────────────────────────────────────────────────

function discountFn(
  name: string,
  check: (x: number, y: number) => boolean,
  f: (x: number, y: number, t: number) => number,
) {
  scalarFn(name, 4, 5, (a) => {
    const v = read(a, "ddnnb")
    if (isError(v)) return v
    const [settle, mat, x, y, basis] = v
    if (settle >= mat || !check(x, y)) return NUM
    const r = f(x, y, yearFrac(settle, mat, basis))
    return Number.isFinite(r) ? r : NUM
  })
}
const positive = (x: number, y: number) => x > 0 && y > 0
discountFn("DISC", positive, (pr, redemption, t) => (1 - pr / redemption) / t)
discountFn("INTRATE", positive, (investment, redemption, t) => (redemption - investment) / investment / t)
discountFn("YIELDDISC", positive, (pr, redemption, t) => (redemption - pr) / pr / t)
discountFn("PRICEDISC", positive, (discount, redemption, t) => redemption * (1 - discount * t))
discountFn("RECEIVED", positive, (investment, discount, t) => {
  const d = 1 - discount * t
  return d <= 0 ? NaN : investment / d
})

/**
 * Years for PRICEMAT and YIELDMAT: 30/360 without YEARFRAC's end-of-February rule for basis 0, and for basis 1 actual
 * days over one year length for all three spans (YEARFRAC's average from issue to settlement).
 */
function maturityYears(issue: number, settle: number, basis: number) {
  const year = basis === 1 ? (settle - issue) / yearFrac(issue, settle, 1) : basis === 3 ? 365 : 360
  return (from: number, to: number) => {
    if (basis === 0) return days360(parts(from), parts(to), false, false) / 360
    if (basis === 4) return days360(parts(from), parts(to), true) / 360
    return (to - from) / year
  }
}

for (const name of ["PRICEMAT", "YIELDMAT"]) {
  scalarFn(name, 5, 6, (a) => {
    const v = read(a, "dddnnb")
    if (isError(v)) return v
    const [settle, mat, issue, rate, x, basis] = v
    if (settle >= mat || rate < 0 || (name === "PRICEMAT" ? x < 0 : x <= 0)) return NUM
    const years = maturityYears(issue, settle, basis)
    const issueMat = years(issue, mat)
    const issueSettle = years(issue, settle)
    const settleMat = years(settle, mat)
    if (name === "PRICEMAT") return (100 * (1 + issueMat * rate)) / (1 + settleMat * x) - 100 * issueSettle * rate
    const paid = x / 100 + issueSettle * rate
    return ((1 + issueMat * rate - paid) / paid) * (1 / settleMat)
  })
}

// ── Treasury bills ───────────────────────────────────────────────────────────────────────────────────────────────

/** Days from settlement to maturity, for a bill of at most a year. */
function billDays(settle: number, mat: number): number | CellError {
  if (settle >= mat || mat > shiftMonths(settle, 12, false)) return NUM
  return mat - settle
}
scalarFn("TBILLPRICE", 3, 3, (a) => {
  const v = read(a, "ddn")
  if (isError(v)) return v
  const dsm = billDays(v[0], v[1])
  if (isError(dsm)) return dsm
  if (v[2] <= 0) return NUM
  const p = 100 * (1 - (v[2] * dsm) / 360)
  return p <= 0 ? NUM : p
})
scalarFn("TBILLYIELD", 3, 3, (a) => {
  const v = read(a, "ddn")
  if (isError(v)) return v
  const dsm = billDays(v[0], v[1])
  if (isError(dsm)) return dsm
  if (v[2] <= 0) return NUM
  return ((100 - v[2]) / v[2]) * (360 / dsm)
})
scalarFn("TBILLEQ", 3, 3, (a) => {
  const v = read(a, "ddn")
  if (isError(v)) return v
  const dsm = billDays(v[0], v[1])
  if (isError(dsm)) return dsm
  const rate = v[2]
  if (rate <= 0) return NUM
  if (dsm <= 182) return (365 * rate) / (360 - rate * dsm)
  // Longer than half a year: the bond-equivalent yield with one coupon in between.
  const pr = 100 * (1 - (rate * dsm) / 360)
  if (pr <= 0) return NUM
  const t = dsm / 365
  return (-2 * t + 2 * Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / pr))) / (2 * t - 1)
})

// ── Odd first and last periods ───────────────────────────────────────────────────────────────────────────────────

/** Quasi-coupon dates from `anchor` stepping by `months` (negative: backwards) past `limit`. */
function quasiDates(anchor: number, months: number, limit: number, eom: boolean): number[] {
  const out = [anchor]
  for (let k = 1; k < 2000; k++) {
    const d = shiftMonths(anchor, months * k, eom)
    out.push(d)
    if (months < 0 ? d <= limit : d >= limit) break
  }
  return out
}

/** Days and period length for the odd-period formulas, by basis (30/360 counts each quasi-period as 360/f). */
function oddDays(start: number, end: number, basis: number): number {
  if (basis === 0) return days360Us(start, end, false)
  if (basis === 4) return days360(parts(start), parts(end), true)
  return end - start
}
function oddLength(pcd: number, ncd: number, basis: number, freq: number): number {
  return basis === 1 ? ncd - pcd : basis === 3 ? 365 / freq : 360 / freq
}

function oddFirstPrice(
  settle: number,
  mat: number,
  issue: number,
  first: number,
  rate: number,
  yld: number,
  redemption: number,
  freq: number,
  basis: number,
): number {
  const months = 12 / freq
  const eom = lastOfMonth(mat)
  const c = coupons(first, mat, freq, basis) as Coupons
  // Regular coupons after the first one.
  const n = coupNum(c) + 1
  const coupon = (100 * rate) / freq
  const base = 1 + yld / freq
  const firstStart = shiftMonths(first, -months, eom)
  if (issue >= firstStart) {
    // Short (or normal) first period.
    const e = oddLength(firstStart, first, basis, freq)
    const dfc = oddDays(issue, first, basis)
    const dsc = oddDays(settle, first, basis)
    const a = oddDays(issue, settle, basis)
    let p = redemption / Math.pow(base, n - 1 + dsc / e)
    p += (coupon * (dfc / e)) / Math.pow(base, dsc / e)
    for (let k = 2; k <= n; k++) p += coupon / Math.pow(base, k - 1 + dsc / e)
    p -= coupon * (a / e)
    return p
  }
  // Long first period: quasi-coupon periods from the first coupon back past the issue.
  const dates = quasiDates(first, -months, issue, eom)
  let dc = 0
  let accrued = 0
  let nq = 0
  let dscNl = 0
  for (let i = 0; i + 1 < dates.length; i++) {
    const ncd = dates[i]
    const pcd = dates[i + 1]
    const nl = oddLength(pcd, ncd, basis, freq)
    const from = Math.max(issue, pcd)
    dc += oddDays(from, ncd, basis) / nl
    if (settle > pcd) {
      accrued += oddDays(from, Math.min(settle, ncd), basis) / nl
    }
    if (settle < ncd && settle >= pcd) dscNl = oddDays(settle, ncd, basis) / nl
    if (settle < pcd) nq++
  }
  let p = redemption / Math.pow(base, n + nq + dscNl - 1)
  p += (coupon * dc) / Math.pow(base, nq + dscNl)
  for (let k = 1; k < n; k++) p += coupon / Math.pow(base, k + nq + dscNl)
  p -= coupon * accrued
  return p
}

function oddFirstArgs(a: Value[]): number[] | CellError {
  const v = read(a, "ddddnnnfb")
  if (isError(v)) return v
  const [settle, mat, issue, first, rate, y, redemption] = v
  if (rate < 0 || y < 0 || redemption <= 0) return NUM
  if (!(issue < settle && settle < first && first < mat)) return NUM
  return v
}
scalarFn("ODDFPRICE", 8, 9, (a) => {
  const v = oddFirstArgs(a)
  if (isError(v)) return v
  const [settle, mat, issue, first, rate, yld, redemption, freq, basis] = v
  return oddFirstPrice(settle, mat, issue, first, rate, yld, redemption, freq, basis)
})
scalarFn("ODDFYIELD", 8, 9, (a) => {
  const v = oddFirstArgs(a)
  if (isError(v)) return v
  const [settle, mat, issue, first, rate, pr, redemption, freq, basis] = v
  if (pr <= 0) return NUM
  return solveDecreasing((y) => oddFirstPrice(settle, mat, issue, first, rate, y, redemption, freq, basis), pr)
})

/** The sums Excel's odd-last-period formulas use, over quasi-coupon periods from the last coupon to maturity. */
function oddLastSums(settle: number, mat: number, last: number, freq: number, basis: number) {
  // Excel counts actual/360 and actual/365 as actual/actual here.
  if (basis === 2 || basis === 3) basis = 1
  const months = 12 / freq
  const eom = lastOfMonth(last)
  const dates = quasiDates(last, months, mat, eom)
  let dc = 0
  let accrued = 0
  let dsc = 0
  for (let i = 0; i + 1 < dates.length; i++) {
    const pcd = dates[i]
    const ncd = dates[i + 1]
    const nl = oddLength(pcd, ncd, basis, freq)
    const end = Math.min(ncd, mat)
    dc += oddDays(pcd, end, basis) / nl
    if (settle > pcd) accrued += oddDays(pcd, Math.min(settle, end), basis) / nl
    if (settle < end) dsc += oddDays(Math.max(settle, pcd), end, basis) / nl
  }
  return { dc, accrued, dsc }
}

function oddLastArgs(a: Value[]): number[] | CellError {
  const v = read(a, "dddnnnfb")
  if (isError(v)) return v
  const [settle, mat, last, rate, y, redemption] = v
  if (rate < 0 || y < 0 || redemption <= 0) return NUM
  if (!(last < settle && settle < mat)) return NUM
  return v
}
scalarFn("ODDLPRICE", 7, 8, (a) => {
  const v = oddLastArgs(a)
  if (isError(v)) return v
  const [settle, mat, last, rate, yld, redemption, freq, basis] = v
  const s = oddLastSums(settle, mat, last, freq, basis)
  const coupon = (100 * rate) / freq
  return (redemption + s.dc * coupon) / (1 + (s.dsc * yld) / freq) - s.accrued * coupon
})
scalarFn("ODDLYIELD", 7, 8, (a) => {
  const v = oddLastArgs(a)
  if (isError(v)) return v
  const [settle, mat, last, rate, pr, redemption, freq, basis] = v
  if (pr <= 0) return NUM
  const s = oddLastSums(settle, mat, last, freq, basis)
  const coupon = (100 * rate) / freq
  const paid = pr + s.accrued * coupon
  return ((redemption + s.dc * coupon - paid) / paid) * (freq / s.dsc)
})
