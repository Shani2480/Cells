import { isError, type CellError } from "../../values"
import { NUM, textToNumber, toText, VALUE } from "../coerce"
import {
  dateSerial,
  daysInMonth,
  isLeapYear,
  MAX_DATE_SERIAL,
  parseDateText,
  parseDateTimeText,
  parseTimeText,
  serialToDate,
  serialToTime,
  serialWeekday,
} from "../dates"
import { toMatrix } from "../evaluate"
import type { Scalar, Value } from "../types"
import { int, num } from "./helpers"
import { define, scalarFn } from "./registry"

type Ctx = Parameters<typeof toMatrix>[1]

/** A date argument: a number, or text that reads as a date. Negative or past 9999 is #NUM!. */
function dateArg(v: Value): number | CellError {
  const n = num(v)
  if (isError(n)) return n
  if (n < 0 || n > MAX_DATE_SERIAL + 0.99999) return NUM
  return n
}

define("TODAY", { min: 0, max: 0, volatile: true, call: (_a, ctx) => Math.floor(ctx.now()) })
define("NOW", { min: 0, max: 0, volatile: true, call: (_a, ctx) => ctx.now() })

scalarFn("DATE", 3, 3, (a) => {
  const y = int(a[0])
  const m = int(a[1])
  const d = int(a[2])
  if (isError(y)) return y
  if (isError(m)) return m
  if (isError(d)) return d
  if (y < 0 || y >= 10000) return NUM
  const s = dateSerial(y, m, d)
  return s < 0 || s > MAX_DATE_SERIAL ? NUM : s
})
scalarFn("TIME", 3, 3, (a) => {
  const h = int(a[0])
  const m = int(a[1])
  const s = int(a[2])
  if (isError(h)) return h
  if (isError(m)) return m
  if (isError(s)) return s
  const total = h * 3600 + m * 60 + s
  if (h > 32767 || m > 32767 || s > 32767 || total < 0) return NUM
  return (total % 86_400) / 86_400
})
scalarFn("DATEVALUE", 1, 1, (a, ctx) => {
  const t = toText(a[0] as Scalar)
  if (isError(t)) return t
  if (typeof a[0] !== "string") return VALUE
  const year = serialToDate(Math.floor(ctx.now())).year
  const d = parseDateText(t, year) ?? parseDateTimeText(t, year)
  return d === null || d < 1 ? VALUE : Math.floor(d)
})
scalarFn("TIMEVALUE", 1, 1, (a, ctx) => {
  if (typeof a[0] !== "string") return VALUE
  const t = a[0]
  const time = parseTimeText(t) ?? parseDateTimeText(t, serialToDate(Math.floor(ctx.now())).year)
  return time === null ? VALUE : time - Math.floor(time)
})

const datePart = (f: (d: ReturnType<typeof serialToDate>) => number) => (a: Value[]) => {
  const n = dateArg(a[0])
  return isError(n) ? n : f(serialToDate(n))
}
scalarFn(
  "YEAR",
  1,
  1,
  datePart((d) => d.year),
)
scalarFn(
  "MONTH",
  1,
  1,
  datePart((d) => d.month),
)
scalarFn(
  "DAY",
  1,
  1,
  datePart((d) => d.day),
)
const timePart = (f: (t: ReturnType<typeof serialToTime>) => number) => (a: Value[]) => {
  const n = dateArg(a[0])
  return isError(n) ? n : f(serialToTime(n))
}
scalarFn(
  "HOUR",
  1,
  1,
  timePart((t) => t.hour),
)
scalarFn(
  "MINUTE",
  1,
  1,
  timePart((t) => t.minute),
)
scalarFn(
  "SECOND",
  1,
  1,
  timePart((t) => t.second),
)

scalarFn("WEEKDAY", 1, 2, (a) => {
  const n = dateArg(a[0])
  const type = a.length > 1 && a[1] !== null ? int(a[1]) : 1
  if (isError(n)) return n
  if (isError(type)) return type
  const w = serialWeekday(n) // 0 = Sunday
  switch (type) {
    case 1:
    case 17:
      return w + 1
    case 2:
    case 11:
      return ((w + 6) % 7) + 1
    case 3:
      return (w + 6) % 7
    case 12:
    case 13:
    case 14:
    case 15:
    case 16:
      // 12: Tuesday is 1 … 16: Saturday is 1.
      return ((w - (type - 10) + 7) % 7) + 1
    default:
      return NUM
  }
})
function isoWeek(serial: number): number {
  const w = (serialWeekday(serial) + 6) % 7 // Monday = 0
  const thursday = Math.floor(serial) - w + 3
  const { year } = serialToDate(thursday)
  const jan1 = dateSerial(year, 1, 1)
  return Math.floor((thursday - jan1) / 7) + 1
}
scalarFn("ISOWEEKNUM", 1, 1, (a) => {
  const n = dateArg(a[0])
  return isError(n) ? n : isoWeek(n)
})
scalarFn("WEEKNUM", 1, 2, (a) => {
  const n = dateArg(a[0])
  const type = a.length > 1 && a[1] !== null ? int(a[1]) : 1
  if (isError(n)) return n
  if (isError(type)) return type
  if (type === 21) return isoWeek(n)
  const starts: Record<number, number> = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 }
  const start = starts[type]
  if (start === undefined) return NUM
  const { year } = serialToDate(n)
  const jan1 = dateSerial(year, 1, 1)
  const offset = (serialWeekday(jan1) - start + 7) % 7
  return Math.floor((Math.floor(n) - jan1 + offset) / 7) + 1
})

function addMonths(serial: number, months: number, endOfMonth: boolean): number {
  const d = serialToDate(serial)
  const total = d.year * 12 + (d.month - 1) + months
  const y = Math.floor(total / 12)
  const m = (total % 12) + 1
  const day = endOfMonth ? daysInMonth(y, m) : Math.min(d.day, daysInMonth(y, m))
  return dateSerial(y, m, day)
}
scalarFn("EDATE", 2, 2, (a) => {
  const n = dateArg(a[0])
  const m = int(a[1])
  if (isError(n)) return n
  if (isError(m)) return m
  const r = addMonths(Math.floor(n), m, false)
  return r < 0 || r > MAX_DATE_SERIAL ? NUM : r
})
scalarFn("EOMONTH", 2, 2, (a) => {
  const n = dateArg(a[0])
  const m = int(a[1])
  if (isError(n)) return n
  if (isError(m)) return m
  const r = addMonths(Math.floor(n), m, true)
  return r < 0 || r > MAX_DATE_SERIAL ? NUM : r
})
scalarFn("DAYS", 2, 2, (a) => {
  const end = dateArg(a[0])
  const start = dateArg(a[1])
  if (isError(end)) return end
  if (isError(start)) return start
  return Math.floor(end) - Math.floor(start)
})
scalarFn("DAYS360", 2, 3, (a) => {
  const s = dateArg(a[0])
  const e = dateArg(a[1])
  if (isError(s)) return s
  if (isError(e)) return e
  const european = a.length > 2 && a[2] !== null && a[2] !== false && a[2] !== 0
  return days360(serialToDate(s), serialToDate(e), european)
})
function days360(
  s: { year: number; month: number; day: number },
  e: { year: number; month: number; day: number },
  european: boolean,
) {
  let sd = s.day
  let ed = e.day
  if (european) {
    if (sd === 31) sd = 30
    if (ed === 31) ed = 30
  } else {
    const lastFeb = (d: typeof s) => d.month === 2 && d.day === daysInMonth(d.year, 2)
    if (lastFeb(s)) {
      if (lastFeb(e)) ed = 30
      sd = 30
    }
    if (ed === 31 && sd >= 30) ed = 30
    if (sd === 31) sd = 30
  }
  return (e.year - s.year) * 360 + (e.month - s.month) * 30 + (ed - sd)
}

scalarFn("DATEDIF", 3, 3, (a) => {
  const s = dateArg(a[0])
  const e = dateArg(a[1])
  const unitText = toText(a[2] as Scalar)
  if (isError(s)) return s
  if (isError(e)) return e
  if (isError(unitText)) return unitText
  if (s > e) return NUM
  const a0 = serialToDate(s)
  const b = serialToDate(e)
  let months = (b.year - a0.year) * 12 + (b.month - a0.month)
  if (b.day < a0.day) months--
  switch (unitText.toUpperCase()) {
    case "D":
      return Math.floor(e) - Math.floor(s)
    case "M":
      return months
    case "Y":
      return Math.floor(months / 12)
    case "YM":
      return months % 12
    case "MD": {
      let d = b.day - a0.day
      if (d < 0) {
        const pm = b.month === 1 ? 12 : b.month - 1
        const py = b.month === 1 ? b.year - 1 : b.year
        d += daysInMonth(py, pm)
      }
      return d
    }
    case "YD": {
      let start = dateSerial(b.year, a0.month, a0.day)
      if (start > Math.floor(e)) start = dateSerial(b.year - 1, a0.month, a0.day)
      return Math.floor(e) - start
    }
    default:
      return NUM
  }
})

scalarFn("YEARFRAC", 2, 3, (a) => {
  let s = dateArg(a[0])
  let e = dateArg(a[1])
  const basis = a.length > 2 && a[2] !== null ? int(a[2]) : 0
  if (isError(s)) return s
  if (isError(e)) return e
  if (isError(basis)) return basis
  if (basis < 0 || basis > 4) return NUM
  s = Math.floor(s)
  e = Math.floor(e)
  if (s > e) [s, e] = [e, s]
  const sd = serialToDate(s)
  const ed = serialToDate(e)
  switch (basis) {
    case 0:
      return days360Us(sd, ed) / 360
    case 1: {
      // Actual/actual: average year length over the years spanned.
      if (
        sd.year === ed.year ||
        (ed.year === sd.year + 1 && (ed.month < sd.month || (ed.month === sd.month && ed.day <= sd.day)))
      ) {
        const leap =
          sd.year === ed.year
            ? isLeapYear(sd.year)
            : (isLeapYear(sd.year) && sd.month < 3) ||
              (isLeapYear(ed.year) && (ed.month > 2 || (ed.month === 2 && ed.day === 29)))
        return (e - s) / (leap ? 366 : 365)
      }
      let days = 0
      for (let y = sd.year; y <= ed.year; y++) days += isLeapYear(y) ? 366 : 365
      return (e - s) / (days / (ed.year - sd.year + 1))
    }
    case 2:
      return (e - s) / 360
    case 3:
      return (e - s) / 365
    default:
      return days360(sd, ed, true) / 360
  }
})
/** The 30/360 US (NASD) day count used by YEARFRAC basis 0. */
function days360Us(s: { year: number; month: number; day: number }, e: { year: number; month: number; day: number }) {
  let sd = s.day
  let ed = e.day
  const lastFeb = (d: typeof s) => d.month === 2 && d.day === daysInMonth(d.year, 2)
  if (lastFeb(s) && lastFeb(e)) ed = 30
  if (lastFeb(s)) sd = 30
  if (ed === 31 && sd >= 30) ed = 30
  if (sd === 31) sd = 30
  return (e.year - s.year) * 360 + (e.month - s.month) * 30 + (ed - sd)
}

// ── Working days ────────────────────────────────────────────────────────────────────────────────────────────────

/** Weekend days as a 7-flag mask, Monday first: a weekend number (1–17) or a string "0000011". */
function weekendMask(v: Scalar | undefined): boolean[] | CellError {
  if (v === undefined || v === null) return [false, false, false, false, false, true, true]
  if (typeof v === "string") {
    if (!/^[01]{7}$/.test(v) || v === "1111111") return VALUE
    return v.split("").map((c) => c === "1")
  }
  if (typeof v !== "number") return VALUE
  const pairs: Record<number, number[]> = {
    1: [5, 6],
    2: [6, 0],
    3: [0, 1],
    4: [1, 2],
    5: [2, 3],
    6: [3, 4],
    7: [4, 5],
  }
  const mask = [false, false, false, false, false, false, false]
  if (pairs[v]) for (const i of pairs[v]) mask[i] = true
  else if (v >= 11 && v <= 17) mask[(v - 11 + 6) % 7] = true
  else return NUM
  return mask
}
function holidaySet(v: Value | undefined, ctx: Ctx): Set<number> | CellError {
  const out = new Set<number>()
  if (v === undefined || v === null) return out
  const m = toMatrix(v, ctx)
  if (isError(m)) return m
  for (const x of m.data) {
    if (x === null) continue
    if (isError(x)) return x
    const n = typeof x === "number" ? x : typeof x === "string" ? textToNumber(x) : null
    if (n === null) return VALUE
    out.add(Math.floor(n))
  }
  return out
}
const isWeekend = (mask: boolean[], serial: number) => mask[(serialWeekday(serial) + 6) % 7]

function networkdays(a: Value[], ctx: Ctx, maskArg: Scalar | undefined, holidays: Value | undefined): Scalar {
  const s = dateArg(a[0])
  const e = dateArg(a[1])
  if (isError(s)) return s
  if (isError(e)) return e
  const mask = weekendMask(maskArg)
  if (isError(mask)) return mask
  const hol = holidaySet(holidays, ctx)
  if (isError(hol)) return hol
  const start = Math.floor(Math.min(s, e))
  const end = Math.floor(Math.max(s, e))
  let count = 0
  for (let d = start; d <= end; d++) if (!isWeekend(mask, d) && !hol.has(d)) count++
  return s <= e ? count : -count
}
define("NETWORKDAYS", {
  min: 2,
  max: 3,
  params: ["scalar", "scalar", "any"],
  call: (a, ctx) => networkdays(a, ctx, undefined, a[2]),
})
define("NETWORKDAYS.INTL", {
  min: 2,
  max: 4,
  params: ["scalar", "scalar", "scalar", "any"],
  call: (a, ctx) => networkdays(a, ctx, a[2] as Scalar, a[3]),
})
function workday(a: Value[], ctx: Ctx, maskArg: Scalar | undefined, holidays: Value | undefined): Scalar {
  const s = dateArg(a[0])
  const n = int(a[1])
  if (isError(s)) return s
  if (isError(n)) return n
  const mask = weekendMask(maskArg)
  if (isError(mask)) return mask
  const hol = holidaySet(holidays, ctx)
  if (isError(hol)) return hol
  let d = Math.floor(s)
  let left = Math.abs(n)
  const step = n < 0 ? -1 : 1
  while (left > 0) {
    d += step
    if (d < 0 || d > MAX_DATE_SERIAL) return NUM
    if (!isWeekend(mask, d) && !hol.has(d)) left--
  }
  return d
}
define("WORKDAY", {
  min: 2,
  max: 3,
  params: ["scalar", "scalar", "any"],
  call: (a, ctx) => workday(a, ctx, undefined, a[2]),
})
define("WORKDAY.INTL", {
  min: 2,
  max: 4,
  params: ["scalar", "scalar", "scalar", "any"],
  call: (a, ctx) => workday(a, ctx, a[2] as Scalar, a[3]),
})
