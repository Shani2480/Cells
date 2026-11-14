/**
 * Excel's 1900 date system: serial 1 is 1900-01-01, and serial 60 is the 29 February 1900 that never existed
 * (kept from Lotus 1-2-3), so serials from 61 on are one more than the days since 1899-12-31. Times are the
 * fraction of a day.
 */

const DAY_MS = 86_400_000
/** 1899-12-30: serial 0 for every date from 1900-03-01 on. */
const EPOCH = Date.UTC(1899, 11, 30)

/**
 * The serial of a date. Months and days roll over as in DATE (month 13 is January next year, day 0 the day before
 * the 1st), counting the fake 29 February 1900: DATE(1900,3,0) is 60.
 */
export function dateSerial(year: number, month: number, day: number): number {
  // DATE treats years 0–1899 as 1900–3799.
  if (year < 1900) year += 1900
  year += Math.floor((month - 1) / 12)
  month = ((((month - 1) % 12) + 12) % 12) + 1
  const days = Math.round((Date.UTC(year, month - 1, 1) - EPOCH) / DAY_MS)
  const first = year === 1900 && month <= 2 ? days - 1 : days
  return first + day - 1
}

export type DateParts = { year: number; month: number; day: number }

/** Year, month and day of a serial (the integer part). Serial 0 is 1900-01-00 and 60 is 1900-02-29, as in Excel. */
export function serialToDate(serial: number): DateParts {
  const whole = Math.floor(serial)
  if (whole === 60) return { year: 1900, month: 2, day: 29 }
  if (whole === 0) return { year: 1900, month: 1, day: 0 }
  const days = whole < 60 ? whole + 1 : whole
  const date = new Date(EPOCH + days * DAY_MS)
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
}

/** 0 = Sunday … 6 = Saturday. Excel thinks 1900-01-01 (serial 1) was a Sunday, so serials before 60 shift by one. */
export function serialWeekday(serial: number): number {
  const whole = Math.floor(serial)
  return (((whole - 1) % 7) + 7) % 7
}

/** Hours, minutes, seconds (rounded to the nearest second, as HOUR/MINUTE/SECOND do). */
export function serialToTime(serial: number) {
  const fraction = serial - Math.floor(serial)
  let seconds = Math.round(fraction * 86_400)
  if (seconds >= 86_400) seconds = 0
  return { hour: Math.floor(seconds / 3600), minute: Math.floor((seconds % 3600) / 60), second: seconds % 60 }
}

/** The latest date Excel accepts: 9999-12-31. */
export const MAX_DATE_SERIAL = 2_958_465

export function isLeapYear(year: number) {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
}

export function daysInMonth(year: number, month: number) {
  return [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]
}

/** A JavaScript time (ms since 1970, local wall clock passed in by the caller) as a serial. */
export function serialFromLocal(year: number, month: number, day: number, h: number, m: number, s: number) {
  return dateSerial(year, month, day) + (h * 3600 + m * 60 + s) / 86_400
}

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]

function monthIndex(word: string): number {
  const w = word.toLowerCase()
  if (w.length < 3) return -1
  const i = MONTHS.findIndex((m) => w.startsWith(m))
  if (i < 0) return -1
  const full = [
    "january",
    "february",
    "march",
    "april",
    "may",
    "june",
    "july",
    "august",
    "september",
    "october",
    "november",
    "december",
  ][i]
  return full.startsWith(w) ? i + 1 : -1
}

/** Years typed with two digits: 00–29 are 2000–2029, 30–99 are 1930–1999 (Excel's default cut-off). */
function fullYear(text: string): number {
  const n = Number(text)
  if (text.length <= 2) return n < 30 ? 2000 + n : 1900 + n
  return n
}

function validDate(year: number, month: number, day: number): number | null {
  if (month < 1 || month > 12 || day < 1 || year < 1900 || year > 9999) return null
  if (year === 1900 && month === 2 && day === 29) return 60
  if (day > daysInMonth(year, month)) return null
  return dateSerial(year, month, day)
}

/** Parses a time of day: 13:30, 1:30 PM, 13:30:15, 1:30:15.5 am, 1 PM. Returns the day fraction. */
export function parseTimeText(text: string): number | null {
  const m = /^(\d{1,4})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d*)?))?\s*([ap]m?)?$/i.exec(text.trim())
  if (!m) return null
  if (m[2] === undefined && !m[4]) return null
  let hour = Number(m[1])
  const minute = m[2] === undefined ? 0 : Number(m[2])
  const second = m[3] === undefined ? 0 : Number(m[3])
  if (minute > 59 || second >= 60) return null
  if (m[4]) {
    if (hour < 1 || hour > 12) return null
    const pm = m[4].toLowerCase().startsWith("p")
    hour = (hour % 12) + (pm ? 12 : 0)
  }
  // Without AM/PM, Excel accepts hours past 24 (as a duration) in values like "25:00".
  return (hour * 3600 + minute * 60 + second) / 86_400
}

/**
 * Parses a date the way Excel (en-US) does on entry and in DATEVALUE: 1/15/2020, 2020-01-15, 15-Jan-2020,
 * Jan 15, 2020, January 2020, 15 Jan. A year left out means the current year (passed in). Returns the serial.
 */
export function parseDateText(text: string, currentYear: number): number | null {
  const t = text.trim()
  let m = /^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{1,4}))?$/.exec(t)
  if (m) return validDate(m[3] ? fullYear(m[3]) : currentYear, Number(m[1]), Number(m[2]))
  m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(t)
  if (m) return validDate(Number(m[1]), Number(m[2]), Number(m[3]))
  m = /^(\d{1,2})[-\s/]([A-Za-z]+)\.?(?:[-\s/,]+(\d{1,4}))?$/.exec(t)
  if (m) {
    const month = monthIndex(m[2])
    if (month > 0) return validDate(m[3] ? fullYear(m[3]) : currentYear, month, Number(m[1]))
  }
  m = /^([A-Za-z]+)\.?[-\s/]+(\d{1,2})(?:(?:,\s*|[-\s/]+)(\d{4}|\d{2}))?$/.exec(t)
  if (m) {
    const month = monthIndex(m[1])
    // "Jan 20" is 20 January this year; "Jan 2020" (4 digits) is 1 January 2020.
    if (month > 0) return validDate(m[3] ? fullYear(m[3]) : currentYear, month, Number(m[2]))
  }
  m = /^([A-Za-z]+)\.?[-\s/]+(\d{4})$/.exec(t)
  if (m) {
    const month = monthIndex(m[1])
    if (month > 0) return validDate(Number(m[2]), month, 1)
  }
  return null
}

/** A date, a time, or a date followed by a time. */
export function parseDateTimeText(text: string, currentYear: number): number | null {
  const t = text.trim()
  const date = parseDateText(t, currentYear)
  if (date !== null) return date
  const time = parseTimeText(t)
  if (time !== null) return time
  const split = /^(.*\S)\s+(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d*)?)?\s*(?:[ap]m?)?|\d{1,2}\s*[ap]m?)$/i.exec(t)
  if (split) {
    const d = parseDateText(split[1], currentYear)
    const tm = parseTimeText(split[2])
    if (d !== null && tm !== null) return d + tm
  }
  return null
}
