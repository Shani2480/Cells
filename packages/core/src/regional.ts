/**
 * How the person writes dates and numbers (File › Options › Advanced; shared with Inkspire). Excel takes these from
 * the system's region; Cells from its settings. They change what typed text means and how numbers and dates show.
 * They never change what is stored: number format codes and formulas stay in Excel's en-US form, so files and formulas
 * mean the same everywhere.
 *
 * One setting for the whole app, as in Excel. Code that must keep en-US rules whatever the setting (reading and
 * writing CSV, the TEXT function, the calculation worker) runs under `withRegional(EN_US, …)` or passes EN_US.
 */

export type DateOrder = "mdy" | "dmy" | "ymd"

export type Regional = {
  decimal: string
  /** "" means no thousands separator is shown or read. */
  group: string
  /** What `/` in a date format shows, and what typed dates are split at (besides `/` and `-`). */
  dateSeparator: string
  dateOrder: DateOrder
  /** The short date as a format code with `/` for the separator ("dd/mm/yyyy"): Excel's built-in format 14. */
  shortDate: string
}

export const EN_US: Regional = { decimal: ".", group: ",", dateSeparator: "/", dateOrder: "mdy", shortDate: "m/d/yyyy" }

let current: Regional = EN_US

export function regional(): Regional {
  return current
}

export function setRegional(next: Regional): void {
  current = next
}

export function isEnUs(r: Regional): boolean {
  return (
    r.decimal === "." &&
    r.group === "," &&
    r.dateSeparator === "/" &&
    r.dateOrder === "mdy" &&
    r.shortDate === "m/d/yyyy"
  )
}

/** Runs `fn` with another setting (en-US for files and functions), then puts the person's back. */
export function withRegional<T>(r: Regional, fn: () => T): T {
  const saved = current
  current = r
  try {
    return fn()
  } finally {
    current = saved
  }
}

/**
 * The setting from the shared settings' short date pattern (`dd/MM/yyyy`, `yyyy-MM-dd`, `d.M.yyyy`) and separators.
 * Blank values keep en-US's.
 */
export function regionalFrom(settings: {
  dateFormat?: string
  decimalSeparator?: string
  groupSeparator?: string
}): Regional {
  const decimal = settings.decimalSeparator || "."
  // A blank thousands separator is a choice only when the decimal one is set too (a region with no grouping).
  const group = settings.decimalSeparator ? (settings.groupSeparator ?? "") : settings.groupSeparator || ","
  const pattern = settings.dateFormat || "M/d/yyyy"
  const fields = [...pattern.matchAll(/y+|M+|d+/g)]
  const order = fields.map((m) => m[0][0].toLowerCase()).join("")
  const dateOrder: DateOrder = order === "dmy" || order === "ymd" ? order : "mdy"
  const separator = /[^yMd]/.exec(pattern)?.[0] ?? "/"
  const shortDate = fields.map((m) => (m[0][0] === "M" ? m[0].toLowerCase() : m[0])).join("/")
  return {
    decimal,
    group: group === decimal ? "" : group,
    dateSeparator: separator,
    dateOrder,
    shortDate: fields.length === 3 ? shortDate : "m/d/yyyy",
  }
}

/** A number written with en-US separators ("1,234.5") in the setting's ("1.234,5"). */
export function localizeNumber(text: string, r: Regional = current): string {
  if (r.decimal === "." && r.group === ",") return text
  return text.replace(/[.,]/g, (c) => (c === "." ? r.decimal : r.group))
}

/**
 * Typed text in the setting's separators ("1.234,5", "$1.000,50") with en-US ones ("1,234.5"), so en-US rules can
 * read it. Text that isn't made of number characters comes back as it was; null when it has a "." or "," the setting
 * doesn't use (so "1.5" with decimal "," and group " " is no number, and may be a date).
 */
export function usSeparators(text: string, r: Regional = current): string | null {
  if (r.decimal === "." && r.group === ",") return text
  if (!/^[+\-\d\s.,eE%$()\u00a0\u202f]*$/.test(text) || !/\d/.test(text)) return text
  const spaceGroup = r.group !== "" && r.group.trim() === ""
  let out = ""
  for (const c of text) {
    if (c === r.decimal) out += "."
    else if (r.group && (c === r.group || (spaceGroup && /[ \u00a0\u202f]/.test(c)))) out += ","
    else if (c === "." || c === ",") return null
    else out += c
  }
  return out
}
