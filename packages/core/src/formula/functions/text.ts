import { isError, type CellError } from "../../values"
import { formatWithCode } from "../../numfmt"
import { NA, numberToText, textToNumber, toBool, toNumber, toText, VALUE } from "../coerce"
import { toMatrix } from "../evaluate"
import { isMatrix, isRef, matrix, type Scalar, type Value } from "../types"
import { wildcardPattern } from "./criteria"
import { forEachScalar, int, num } from "./helpers"
import { define, scalarFn } from "./registry"

const text = (v: Value) => toText(v as Scalar)

scalarFn("LEN", 1, 1, (a) => {
  const s = text(a[0])
  return isError(s) ? s : s.length
})
/**
 * s.slice(from, to) without splitting a surrogate pair: Excel counts UTF-16 units (LEN("😀") is 2) but LEFT("😀x",1)
 * still returns the whole emoji.
 */
function slicePairs(s: string, from: number, to = s.length) {
  const high = (i: number) => /[\uD800-\uDBFF]/.test(s[i] ?? "")
  const low = (i: number) => /[\uDC00-\uDFFF]/.test(s[i] ?? "")
  if (from > 0 && low(from) && high(from - 1)) from--
  if (to > 0 && to < s.length && high(to - 1) && low(to)) to++
  return s.slice(from, to)
}

scalarFn("LEFT", 1, 2, (a) => {
  const s = text(a[0])
  const n = a.length > 1 ? int(a[1]) : 1
  if (isError(s)) return s
  if (isError(n)) return n
  if (n < 0) return VALUE
  return slicePairs(s, 0, n)
})
scalarFn("RIGHT", 1, 2, (a) => {
  const s = text(a[0])
  const n = a.length > 1 ? int(a[1]) : 1
  if (isError(s)) return s
  if (isError(n)) return n
  if (n < 0) return VALUE
  return n === 0 ? "" : slicePairs(s, Math.max(0, s.length - n))
})
scalarFn("MID", 3, 3, (a) => {
  const s = text(a[0])
  const start = int(a[1])
  const n = int(a[2])
  if (isError(s)) return s
  if (isError(start)) return start
  if (isError(n)) return n
  if (start < 1 || n < 0) return VALUE
  return slicePairs(s, Math.min(s.length, start - 1), start - 1 + n)
})
scalarFn("UPPER", 1, 1, (a) => {
  const s = text(a[0])
  return isError(s) ? s : s.toUpperCase()
})
scalarFn("LOWER", 1, 1, (a) => {
  const s = text(a[0])
  return isError(s) ? s : s.toLowerCase()
})
scalarFn("PROPER", 1, 1, (a) => {
  const s = text(a[0])
  if (isError(s)) return s
  // A letter after a non-letter is capitalised; every other letter is lower case.
  return s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, p, c) => p + c.toUpperCase())
})
scalarFn("TRIM", 1, 1, (a) => {
  const s = text(a[0])
  return isError(s) ? s : s.replace(/ +/g, " ").replace(/^ | $/g, "")
})
scalarFn("CLEAN", 1, 1, (a) => {
  const s = text(a[0])
  // eslint-disable-next-line no-control-regex
  return isError(s) ? s : s.replace(/[\x00-\x1f]/g, "")
})
scalarFn("REPT", 2, 2, (a) => {
  const s = text(a[0])
  const n = int(a[1])
  if (isError(s)) return s
  if (isError(n)) return n
  if (n < 0 || s.length * n > 32767) return VALUE
  return s.repeat(n)
})
scalarFn("EXACT", 2, 2, (a) => {
  const x = text(a[0])
  const y = text(a[1])
  if (isError(x)) return x
  if (isError(y)) return y
  return x === y
})
scalarFn("CHAR", 1, 1, (a) => {
  const n = int(a[0])
  if (isError(n)) return n
  if (n < 1 || n > 255) return VALUE
  // Windows-1252 for 128–159.
  const cp1252 = [
    8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141, 381, 143, 144, 8216, 8217, 8220, 8221,
    8226, 8211, 8212, 732, 8482, 353, 8250, 339, 157, 382, 376,
  ]
  return String.fromCharCode(n >= 128 && n < 160 ? cp1252[n - 128] : n)
})
scalarFn("CODE", 1, 1, (a) => {
  const s = text(a[0])
  if (isError(s)) return s
  if (s === "") return VALUE
  const c = s.charCodeAt(0)
  const cp1252: Record<number, number> = {
    8364: 128,
    8218: 130,
    402: 131,
    8222: 132,
    8230: 133,
    8224: 134,
    8225: 135,
    710: 136,
    8240: 137,
    352: 138,
    8249: 139,
    338: 140,
    381: 142,
    8216: 145,
    8217: 146,
    8220: 147,
    8221: 148,
    8226: 149,
    8211: 150,
    8212: 151,
    732: 152,
    8482: 153,
    353: 154,
    8250: 155,
    339: 156,
    382: 158,
    376: 159,
  }
  return c < 256 ? c : (cp1252[c] ?? 63)
})
scalarFn("UNICHAR", 1, 1, (a) => {
  const n = int(a[0])
  if (isError(n)) return n
  if (n < 1 || n > 0x10ffff) return VALUE
  return String.fromCodePoint(n)
})
scalarFn("UNICODE", 1, 1, (a) => {
  const s = text(a[0])
  if (isError(s)) return s
  if (s === "") return VALUE
  return s.codePointAt(0)!
})

function findIn(caseSensitive: boolean, wildcards: boolean) {
  return (a: Value[]) => {
    const needle = text(a[0])
    const hay = text(a[1])
    const start = a.length > 2 && a[2] !== null ? int(a[2]) : 1
    if (isError(needle)) return needle
    if (isError(hay)) return hay
    if (isError(start)) return start
    if (start < 1 || start > hay.length + 1) return VALUE
    if (needle === "") return start
    if (wildcards) {
      const pattern = wildcardPattern(needle)
      if (pattern) {
        const body = pattern.source.slice(1, -1)
        const re = new RegExp(body, "i")
        const m = re.exec(hay.slice(start - 1))
        return m ? m.index + start : VALUE
      }
    }
    const i = caseSensitive
      ? hay.indexOf(needle, start - 1)
      : hay.toLowerCase().indexOf(needle.toLowerCase(), start - 1)
    return i < 0 ? VALUE : i + 1
  }
}
scalarFn("FIND", 2, 3, findIn(true, false))
scalarFn("SEARCH", 2, 3, findIn(false, true))
scalarFn("SUBSTITUTE", 3, 4, (a) => {
  const s = text(a[0])
  const from = text(a[1])
  const to = text(a[2])
  const nth = a.length > 3 && a[3] !== null ? int(a[3]) : null
  if (isError(s)) return s
  if (isError(from)) return from
  if (isError(to)) return to
  if (nth !== null && isError(nth)) return nth
  if (from === "") return s
  if (nth === null) return s.split(from).join(to)
  if (nth < 1) return VALUE
  let index = -1
  for (let k = 0; k < nth; k++) {
    index = s.indexOf(from, index + 1)
    if (index < 0) return s
  }
  return s.slice(0, index) + to + s.slice(index + from.length)
})
scalarFn("REPLACE", 4, 4, (a) => {
  const s = text(a[0])
  const start = int(a[1])
  const n = int(a[2])
  const r = text(a[3])
  if (isError(s)) return s
  if (isError(start)) return start
  if (isError(n)) return n
  if (isError(r)) return r
  if (start < 1 || n < 0) return VALUE
  return s.slice(0, start - 1) + r + s.slice(start - 1 + n)
})

define("CONCAT", {
  min: 1,
  max: 253,
  call: (args, ctx) => {
    let out = ""
    let error: CellError | null = null
    for (const arg of args) {
      if (isRef(arg) || isMatrix(arg)) {
        const m = isMatrix(arg) ? arg : toMatrix(arg, ctx)
        if (isError(m)) return m
        for (const v of m.data) {
          const t = toText(v)
          if (isError(t)) return t
          out += t
        }
        continue
      }
      const t = toText(arg)
      if (isError(t)) {
        error = t
        break
      }
      out += t
    }
    if (error) return error
    return out.length > 32767 ? VALUE : out
  },
})
define("CONCATENATE", {
  min: 1,
  max: 255,
  params: ["scalar"],
  call: (args) => {
    let out = ""
    for (const arg of args) {
      const t = toText(arg as Scalar)
      if (isError(t)) return t
      out += t
    }
    return out
  },
})
define("TEXTJOIN", {
  min: 3,
  max: 252,
  call: (args, ctx) => {
    const delimiters = toMatrix(args[0], ctx)
    if (isError(delimiters)) return delimiters
    const delims = delimiters.data.map((d) => toText(d))
    for (const d of delims) if (isError(d)) return d
    const ignore = toBool(scalar(args[1], ctx))
    if (isError(ignore)) return ignore
    const parts: string[] = []
    let error: CellError | null = null
    const rest = args.slice(2)
    for (const arg of rest) {
      const m = isRef(arg) || isMatrix(arg) ? toMatrix(arg, ctx) : null
      const values: Scalar[] = m ? (isError(m) ? [m] : m.data) : [arg as Scalar]
      for (const v of values) {
        const t = toText(v)
        if (isError(t)) {
          error = t
          break
        }
        if (ignore && t === "") continue
        parts.push(t)
      }
      if (error) return error
    }
    let out = ""
    parts.forEach((p, i) => {
      out += (i > 0 ? (delims[(i - 1) % delims.length] as string) : "") + p
    })
    return out.length > 32767 ? VALUE : out
  },
})
function scalar(v: Value, ctx: Parameters<typeof toMatrix>[1]): Scalar {
  const m = toMatrix(v, ctx)
  return isError(m) ? m : m.data[0]
}

scalarFn("TEXT", 2, 2, (a) => {
  const v = a[0] as Scalar
  const code = text(a[1])
  if (isError(code)) return code
  if (v === null) return formatWithCode(0, code).text
  if (typeof v === "string") {
    const n = textToNumber(v)
    return formatWithCode(n ?? v, code).text
  }
  return formatWithCode(v as number | boolean, code).text
})
scalarFn("VALUE", 1, 1, (a) => {
  const v = a[0] as Scalar
  if (typeof v === "number") return v
  if (v === null) return 0
  if (typeof v === "boolean") return VALUE
  return textToNumber(v as string) ?? VALUE
})
scalarFn("NUMBERVALUE", 1, 3, (a) => {
  const s = text(a[0])
  if (isError(s)) return s
  const dec = a.length > 1 && a[1] !== null ? text(a[1]) : "."
  const group = a.length > 2 && a[2] !== null ? text(a[2]) : ","
  if (isError(dec)) return dec
  if (isError(group)) return group
  let t = s.replace(/\s+/g, "")
  if (t === "") return 0
  t = t.split(group[0] ?? ",").join("")
  t = t.replace(dec[0] ?? ".", ".")
  let percent = 0
  while (t.endsWith("%")) {
    percent++
    t = t.slice(0, -1)
  }
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(t)) return VALUE
  return Number(t) / 100 ** percent
})
scalarFn("FIXED", 1, 3, (a) => {
  const x = num(a[0])
  const d = a.length > 1 && a[1] !== null ? int(a[1]) : 2
  const noCommas = a.length > 2 ? toBool(a[2] as Scalar) : false
  if (isError(x)) return x
  if (isError(d)) return d
  if (isError(noCommas)) return noCommas
  const digits = Math.max(0, d)
  const code = `${noCommas ? "0" : "#,##0"}${digits ? "." + "0".repeat(digits) : ""}`
  const rounded = d < 0 ? Math.round(x / 10 ** -d) * 10 ** -d : x
  return formatWithCode(rounded, code).text
})
scalarFn("DOLLAR", 1, 2, (a) => {
  const x = num(a[0])
  const d = a.length > 1 && a[1] !== null ? int(a[1]) : 2
  if (isError(x)) return x
  if (isError(d)) return d
  const digits = Math.max(0, d)
  return formatWithCode(
    x,
    `$#,##0${digits ? "." + "0".repeat(digits) : ""};($#,##0${digits ? "." + "0".repeat(digits) : ""})`,
  ).text
})
define("T", {
  min: 1,
  max: 1,
  params: ["scalar"],
  call: (a) => (typeof a[0] === "string" ? a[0] : ""),
})

// TEXTBEFORE / TEXTAFTER / TEXTSPLIT (Excel 365).
function delimiterList(v: Value, ctx: Parameters<typeof toMatrix>[1]): string[] | CellError {
  const m = toMatrix(v, ctx)
  if (isError(m)) return m
  const out: string[] = []
  for (const d of m.data) {
    const t = toText(d)
    if (isError(t)) return t
    out.push(t)
  }
  return out
}
function findDelimiter(
  s: string,
  delims: string[],
  from: number,
  caseInsensitive: boolean,
): { index: number; length: number } | null {
  const hay = caseInsensitive ? s.toLowerCase() : s
  let best: { index: number; length: number } | null = null
  for (const d of delims) {
    if (d === "") continue
    const i = hay.indexOf(caseInsensitive ? d.toLowerCase() : d, from)
    if (i >= 0 && (!best || i < best.index)) best = { index: i, length: d.length }
  }
  return best
}
function textAround(after: boolean) {
  return (a: Value[], ctx: Parameters<typeof toMatrix>[1]): Value => {
    const s = toText(scalar(a[0], ctx))
    if (isError(s)) return s
    const delims = delimiterList(a[1], ctx)
    if (isError(delims)) return delims
    const nth = a.length > 2 && a[2] !== null ? int(a[2]) : 1
    const mode = a.length > 3 && a[3] !== null ? int(a[3]) : 0
    const matchEnd = a.length > 4 && a[4] !== null ? toBool(scalar(a[4], ctx)) : false
    for (const v of [nth, mode, matchEnd]) if (isError(v)) return v
    const n = nth as number
    if (n === 0 || Math.abs(n) > s.length + 1) return VALUE
    const ci = mode === 1
    // All delimiter positions, left to right.
    const hits: { index: number; length: number }[] = []
    let from = 0
    for (;;) {
      const h = findDelimiter(s, delims, from, ci)
      if (!h) break
      hits.push(h)
      from = h.index + Math.max(1, h.length)
    }
    if (matchEnd) {
      if (n > 0) hits.push({ index: s.length, length: 0 })
      else hits.unshift({ index: 0, length: 0 })
    }
    const pick = n > 0 ? hits[n - 1] : hits[hits.length + n]
    if (!pick) {
      const fallback = a.length > 5 ? scalar(a[5], ctx) : NA
      return fallback
    }
    return after ? s.slice(pick.index + pick.length) : s.slice(0, pick.index)
  }
}
define("TEXTBEFORE", { min: 2, max: 6, call: textAround(false) })
define("TEXTAFTER", { min: 2, max: 6, call: textAround(true) })
define("TEXTSPLIT", {
  min: 2,
  max: 6,
  call: (a, ctx) => {
    const s = toText(scalar(a[0], ctx))
    if (isError(s)) return s
    const colDelims = a[1] === null ? [] : delimiterList(a[1], ctx)
    if (isError(colDelims)) return colDelims
    const rowDelims = a.length > 2 && a[2] !== null ? delimiterList(a[2], ctx) : []
    if (isError(rowDelims)) return rowDelims
    const ignoreEmpty = a.length > 3 && a[3] !== null ? toBool(scalar(a[3], ctx)) : false
    if (isError(ignoreEmpty)) return ignoreEmpty
    const ci = a.length > 4 && a[4] !== null ? scalar(a[4], ctx) === 1 : false
    const pad = a.length > 5 ? scalar(a[5], ctx) : NA
    const split = (t: string, delims: string[]) => {
      if (!delims.length) return [t]
      const out: string[] = []
      let from = 0
      for (;;) {
        const h = findDelimiter(t, delims, from, ci)
        if (!h) {
          out.push(t.slice(from))
          break
        }
        out.push(t.slice(from, h.index))
        from = h.index + h.length
      }
      return ignoreEmpty ? out.filter((x) => x !== "") : out
    }
    const rows = split(s, rowDelims).map((r) => split(r, colDelims))
    const width = Math.max(...rows.map((r) => r.length))
    const data: Scalar[] = []
    for (const r of rows) for (let c = 0; c < width; c++) data.push(c < r.length ? r[c] : pad)
    return matrix(rows.length, width, data)
  },
})

export { forEachScalar, numberToText, toNumber }
