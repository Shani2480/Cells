/**
 * What Excel (en-US) makes of typed dates, times, currency, fractions and scientific numbers: the value and the
 * number format it sets on entry. The rules and every example in the tests were read from Excel 16 by typing the same
 * text into a cell (Range.Formula, which parses as typing does).
 */

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
const FULL_MONTHS = [
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
]

export type Entry = { value: number; numberFormat?: string }

/** A typed date's serial (null when the day does not exist): days from Excel's 1900 epoch (serial 1 = 1900-01-01, with Excel's phantom 29 Feb 1900). */
export function validDateSerial(year: number, month: number, day: number): number | null {
  if (month < 1 || month > 12 || day < 1) return null
  const days = new Date(Date.UTC(year, month, 0)).getUTCDate()
  if (day > days) return null
  if (year < 1900 || year > 9999) return null
  const ms = Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)
  const serial = Math.round(ms / 86_400_000)
  // Before 1 March 1900 Excel counts one day less (it believes 1900 was a leap year).
  return serial < 61 ? serial - 1 : serial
}

function monthOf(word: string): number | null {
  const w = word.toLowerCase()
  if (w.length < 3) return null
  const full = FULL_MONTHS.findIndex((m) => m === w)
  if (full >= 0) return full + 1
  const short = MONTHS.findIndex((m) => m === w || (w.length === 4 && w === "sept" && m === "sep"))
  return short >= 0 ? short + 1 : null
}

/** Two-digit years: 00–29 are 2000s, 30–99 are 1900s (Excel's default cut-off). */
function fullYear(text: string): number {
  const n = Number(text)
  if (text.length > 2) return n
  return n < 30 ? 2000 + n : 1900 + n
}

/** The date part: its serial and format, or null when the text is not a date Excel recognises. */
function parseDate(text: string, thisYear: number): Entry | null {
  const t = text.trim()
  let m: RegExpExecArray | null
  // 2026-01-15, 2026/5/15
  if ((m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(t))) {
    const v = validDateSerial(Number(m[1]), Number(m[2]), Number(m[3]))
    return v === null ? null : { value: v, numberFormat: "m/d/yyyy" }
  }
  // 1/15/2026, 1/15/26, 1/2/3 (and with dashes)
  if ((m = /^(\d{1,2})([-/])(\d{1,2})\2(\d{1,4})$/.exec(t))) {
    const v = validDateSerial(fullYear(m[4]), Number(m[1]), Number(m[3]))
    return v === null ? null : { value: v, numberFormat: "m/d/yyyy" }
  }
  // 1/15 (this year), 5/2026 (month and year)
  if ((m = /^(\d{1,2})\/(\d{1,4})$/.exec(t))) {
    const a = Number(m[1])
    const b = Number(m[2])
    if (m[2].length <= 2 && b <= 31) {
      const v = validDateSerial(thisYear, a, b)
      if (v !== null) return { value: v, numberFormat: "d-mmm" }
    }
    if (m[2].length === 4 || b > 31) {
      const v = validDateSerial(fullYear(m[2]), a, 1)
      return v === null ? null : { value: v, numberFormat: "mmm-yy" }
    }
    return null
  }
  const parts = t
    .replace(/,/g, " ")
    .split(/[\s/-]+/)
    .filter(Boolean)
  if (parts.length < 2 || parts.length > 3) return null
  const months = parts.map(monthOf)
  const at = months.findIndex((x) => x !== null)
  if (at < 0 || months.filter((x) => x !== null).length > 1) return null
  const month = months[at]!
  const nums = parts.filter((_, i) => i !== at)
  if (!nums.every((p) => /^\d{1,4}$/.test(p))) return null
  if (parts.length === 2) {
    const n = nums[0]
    // 15-Jan, Jan 15, mar-5: a day this year. Jan-2026: a month.
    if (n.length <= 2 && Number(n) <= 31) {
      const v = validDateSerial(thisYear, month, Number(n))
      if (v !== null) return { value: v, numberFormat: "d-mmm" }
      if (at === 0) {
        const y = validDateSerial(fullYear(n), month, 1)
        return y === null ? null : { value: y, numberFormat: "mmm-yy" }
      }
      return null
    }
    const v = validDateSerial(fullYear(n), month, 1)
    return v === null ? null : { value: v, numberFormat: "mmm-yy" }
  }
  // Three parts: 15-Jan-2026, 15 Jan 2026, January 15, 2026. "Jan 15 2026" without the comma stays text.
  if (at === 1) {
    const v = validDateSerial(fullYear(nums[1]), month, Number(nums[0]))
    return v === null ? null : { value: v, numberFormat: "d-mmm-yy" }
  }
  if (at === 0 && /,/.test(t)) {
    const v = validDateSerial(fullYear(nums[1]), month, Number(nums[0]))
    return v === null ? null : { value: v, numberFormat: "d-mmm-yy" }
  }
  return null
}

/** The time part: a fraction of a day and its format. */
function parseTime(text: string): Entry | null {
  const m = /^(\d{1,5}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm|a|p)?$/i.exec(text.trim())
  if (!m) return null
  let h = Number(m[1])
  const min = Number(m[2])
  const sec = m[3] === undefined ? 0 : Number(m[3])
  if (min > 59 || sec >= 60) return null
  const ampm = m[4]?.toLowerCase()
  if (ampm) {
    if (h < 1 || h > 12) return null
    if (h === 12) h = 0
    if (ampm.startsWith("p")) h += 12
  }
  const value = (h * 3600 + min * 60 + sec) / 86400
  let numberFormat: string
  if (!ampm && h >= 24) numberFormat = "[h]:mm:ss"
  else numberFormat = (m[3] === undefined ? "h:mm" : "h:mm:ss") + (ampm ? " AM/PM" : "")
  return { value, numberFormat }
}

/** Dates, times and both together. `thisYear` fills in dates typed without a year. */
export function parseDateTime(text: string, thisYear = new Date().getFullYear()): Entry | null {
  const t = text.trim()
  if (!/\d/.test(t)) return null
  const time = parseTime(t)
  if (time) return time
  // A date followed by a time: "1/15/2026 13:45", "1/15/2026 1:45 PM".
  const split = /^(.*\S)\s+(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?(?:\s*(?:am|pm|a|p))?)$/i.exec(t)
  if (split) {
    const date = parseDate(split[1], thisYear)
    const clock = parseTime(split[2])
    if (date && clock) return { value: date.value + clock.value, numberFormat: "m/d/yyyy h:mm" }
    return null
  }
  return parseDate(t, thisYear)
}

/** "$1,000", "-$5", "$-5", "($5)", "$ 5": the amount and Excel's currency format (two decimals when typed). */
export function parseCurrency(text: string): Entry | null {
  const m = /^(\()?\s*(-)?\s*\$\s*(-)?\s*(\d{1,3}(?:,\d{3})+|\d+)(\.\d*)?\s*(\))?$/.exec(text.trim())
  if (!m) return null
  const open = !!m[1]
  const close = !!m[6]
  if (open !== close) return null
  if ((m[2] && m[3]) || ((m[2] || m[3]) && open)) return null
  const value = Number(m[4].replace(/,/g, "") + (m[5] && m[5] !== "." ? m[5] : ""))
  const negative = open || !!m[2] || !!m[3]
  const decimals = m[5] && m[5].length > 1
  const numberFormat = decimals ? "$#,##0.00_);[Red]($#,##0.00)" : "$#,##0_);[Red]($#,##0)"
  return { value: negative ? -value : value, numberFormat }
}

/** "1 1/2", "0 1/2": a whole number and a fraction (a bare "1/2" is a date). */
export function parseFraction(text: string): Entry | null {
  const m = /^([+-])?(\d+) (\d+)\/(\d+)$/.exec(text.trim())
  if (!m) return null
  const num = Number(m[3])
  const den = Number(m[4])
  if (den === 0 || num >= den * 10 ** 6) return null
  const value = (Number(m[2]) + num / den) * (m[1] === "-" ? -1 : 1)
  return { value, numberFormat: m[4].length > 1 ? "# ??/??" : "# ?/?" }
}
