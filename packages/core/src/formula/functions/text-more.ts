/**
 * The rest of Excel's text functions: REGEXTEST/REGEXEXTRACT/REGEXREPLACE, ARRAYTOTEXT and VALUETOTEXT, ENCODEURL,
 * BAHTTEXT, PHONETIC, and the double-byte ones (ASC, DBCS, LENB, LEFTB…), which in a language without double-byte
 * characters do what their plain twins do, as Excel does there.
 */
import "./text"
import { roundDecimal } from "../../rounding"
import { isError, type CellError } from "../../values"
import { NA, numberToText, toText, VALUE } from "../coerce"
import { toMatrix, toScalar } from "../evaluate"
import { isRef, matrix, type Scalar, type Value } from "../types"
import { int, num } from "./helpers"
import { define, FUNCTIONS, scalarFn } from "./registry"

const text = (v: Value) => toText(v as Scalar)

for (const [twin, of] of [
  ["LENB", "LEN"],
  ["LEFTB", "LEFT"],
  ["RIGHTB", "RIGHT"],
  ["MIDB", "MID"],
  ["FINDB", "FIND"],
  ["SEARCHB", "SEARCH"],
  ["REPLACEB", "REPLACE"],
  ["USDOLLAR", "DOLLAR"],
] as const) {
  const def = FUNCTIONS.get(of)
  if (def) define(twin, def)
}
for (const name of ["ASC", "DBCS"]) scalarFn(name, 1, 1, (a) => text(a[0]))

/** A value as ARRAYTOTEXT/VALUETOTEXT write it: strict puts text in quotes (inner quotes doubled). */
function valueText(v: Scalar, strict: boolean): string {
  if (v === null) return ""
  if (isError(v)) return v.error
  if (typeof v === "string") return strict ? `"${v.replace(/"/g, '""')}"` : v
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE"
  return numberToText(v)
}

function formatArg(v: Value | undefined): boolean | CellError {
  if (v === undefined || v === null) return false
  const f = num(v)
  if (isError(f)) return f
  if (f === 0 || f === 1) return f === 1
  return VALUE
}

define("VALUETOTEXT", {
  min: 1,
  max: 2,
  params: ["scalar"],
  errors: true,
  call: (a) => {
    const strict = formatArg(isError(a[1] as Scalar) ? VALUE : a[1])
    if (isError(strict)) return strict
    return valueText(a[0] as Scalar, strict)
  },
})

define("ARRAYTOTEXT", {
  min: 1,
  max: 2,
  call: (a, ctx) => {
    const strict = formatArg(a[1] === undefined ? undefined : toScalar(a[1], ctx))
    if (isError(strict)) return strict
    const m = toMatrix(a[0], ctx)
    if (isError(m)) return m
    const rows: string[] = []
    for (let r = 0; r < m.rows; r++)
      rows.push(
        m.data
          .slice(r * m.cols, (r + 1) * m.cols)
          .map((v) => valueText(v, strict))
          .join(strict ? "," : ", "),
      )
    return strict ? `{${rows.join(";")}}` : rows.join(", ")
  },
})

scalarFn("ENCODEURL", 1, 1, (a) => {
  const t = text(a[0])
  if (isError(t)) return t
  let out = ""
  // encodeURIComponent gives UTF-8 percent escapes; Excel also escapes ! ' ( ) * ~.
  for (const ch of t) {
    out += /[A-Za-z0-9\-_.]/.test(ch)
      ? ch
      : encodeURIComponent(ch).replace(/[!'()*~]/g, (c) => "%" + c.charCodeAt(0).toString(16).toUpperCase())
  }
  return out
})

// BAHTTEXT: a number in Thai words, in baht and satang.
const DIGITS = ["", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]
const PLACES = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"]

/** 1–999,999 in words: 1 in the tens is just สิบ, 2 there ยี่สิบ, and 1 after higher digits เอ็ด. */
function thaiGroup(n: number): string {
  const digits = String(n).split("").map(Number)
  let out = ""
  digits.forEach((d, i) => {
    const place = digits.length - 1 - i
    if (d === 0) return
    if (place === 0) out += d === 1 && n > 9 ? "เอ็ด" : DIGITS[d]
    else if (place === 1) out += (d === 1 ? "" : d === 2 ? "ยี่" : DIGITS[d]) + PLACES[1]
    else out += DIGITS[d] + PLACES[place]
  })
  return out
}

function thaiWords(n: number): string {
  if (n < 1_000_000) return thaiGroup(n)
  const high = Math.floor(n / 1_000_000)
  const low = n - high * 1_000_000
  return thaiWords(high) + "ล้าน" + (low ? thaiGroup(low) : "")
}

scalarFn("BAHTTEXT", 1, 1, (a) => {
  const x = num(a[0])
  if (isError(x)) return x
  const total = Math.round(roundDecimal(Math.abs(x), 2) * 100)
  const baht = Math.floor(total / 100)
  const satang = total % 100
  if (total === 0) return "ศูนย์บาทถ้วน"
  let out = x < 0 ? "ลบ" : ""
  if (baht) out += thaiWords(baht) + "บาท"
  out += satang ? thaiGroup(satang) + "สตางค์" : "ถ้วน"
  return out
})

/** PHONETIC reads the furigana stored with Japanese text; cells without any (all of ours) give #N/A, as in Excel. */
define("PHONETIC", {
  min: 1,
  max: 1,
  call: (a) => (isRef(a[0]) ? NA : VALUE),
})

// REGEXTEST / REGEXEXTRACT / REGEXREPLACE (Excel 365; JavaScript's regular expressions stand in for PCRE2).
function regex(pattern: Value, caseArg: Value | undefined, global: boolean): RegExp | CellError {
  const p = text(pattern)
  if (isError(p)) return p
  let insensitive = false
  if (caseArg !== undefined && caseArg !== null) {
    const c = int(caseArg)
    if (isError(c)) return c
    if (c !== 0 && c !== 1) return VALUE
    insensitive = c === 1
  }
  try {
    return new RegExp(p, (global ? "g" : "") + (insensitive ? "i" : "") + "u")
  } catch {
    return VALUE
  }
}

scalarFn("REGEXTEST", 2, 3, (a) => {
  const t = text(a[0])
  if (isError(t)) return t
  const re = regex(a[1], a[2], false)
  return isError(re) ? re : re.test(t)
})

scalarFn("REGEXEXTRACT", 2, 4, (a) => {
  const t = text(a[0])
  if (isError(t)) return t
  const mode = a[2] === undefined || a[2] === null ? 0 : int(a[2])
  if (isError(mode)) return mode
  if (mode < 0 || mode > 2) return VALUE
  const re = regex(a[1], a[3], mode === 1)
  if (isError(re)) return re
  if (mode === 1) {
    const all = [...t.matchAll(re)].map((m) => m[0])
    return all.length ? matrix(all.length, 1, all) : NA
  }
  const m = re.exec(t)
  if (!m) return NA
  if (mode === 0) return m[0]
  const groups = m.slice(1).map((g) => g ?? "")
  return groups.length ? matrix(1, groups.length, groups) : m[0]
})

scalarFn("REGEXREPLACE", 3, 5, (a) => {
  const t = text(a[0])
  if (isError(t)) return t
  const replacement = text(a[2])
  if (isError(replacement)) return replacement
  const occurrence = a[3] === undefined || a[3] === null ? 0 : int(a[3])
  if (isError(occurrence)) return occurrence
  const re = regex(a[1], a[4], true)
  if (isError(re)) return re
  if (occurrence === 0) return t.replace(re, replacement)
  const count = [...t.matchAll(re)].length
  const wanted = occurrence > 0 ? occurrence : count + occurrence + 1
  let seen = 0
  return t.replace(re, (match: string, ...rest: unknown[]) => {
    seen++
    if (seen !== wanted) return match
    const groups = rest.slice(
      0,
      rest.findIndex((x) => typeof x === "number"),
    ) as (string | undefined)[]
    return replacement.replace(/\$(\$|&|\d{1,2})/g, (_, k: string) =>
      k === "$" ? "$" : k === "&" ? match : (groups[Number(k) - 1] ?? ""),
    )
  })
})
