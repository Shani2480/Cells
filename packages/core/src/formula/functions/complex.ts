/**
 * Complex numbers as text ("3+4i", "2-j"): COMPLEX and the IM* functions, as Excel's Analysis ToolPak reads and
 * writes them. Results are written with 15 significant digits, in full up to 21 characters and in E notation beyond.
 * Excel works in polar form for powers and roots, and its sine and cosine of a multiple of π/2 come out a little
 * off zero (IMPOWER("i",2) is -1+1.22514845490862E-16i there); `sin`/`cos` below reproduce that.
 */
import { isError, type CellError } from "../../values"
import { NUM, numberToText, VALUE, DIV0 } from "../coerce"
import type { Scalar, Value } from "../types"
import { atpNumber } from "./engineering"
import { forEachScalar } from "./helpers"
import { define, scalarFn } from "./registry"

type Suffix = "i" | "j"
type Complex = { re: number; im: number; suffix: Suffix | null }

const NUMBER = String.raw`(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?`

/** A number as Excel reads typed digits: anything past the 15th significant digit counts as 0. */
function readNumber(text: string): number {
  const m = /^([+-]?)(\d*)\.?(\d*)((?:[eE][+-]?\d+)?)$/.exec(text)
  if (!m) return Number(text)
  const digits = m[2] + m[3]
  const first = digits.search(/[1-9]/)
  if (first < 0 || digits.length - first <= 15) return Number(text)
  const kept = digits.slice(0, first + 15) + "0".repeat(digits.length - first - 15)
  return Number(`${m[1]}${kept.slice(0, m[2].length)}.${kept.slice(m[2].length)}${m[4]}`)
}
const REAL = new RegExp(`^([+-]?${NUMBER})$`)
const IMAG = new RegExp(`^([+-]?)(${NUMBER})?([ij])$`)
const BOTH = new RegExp(`^([+-]?${NUMBER})([+-])([+-]?)(${NUMBER})?([ij])$`)

/** A complex argument: text in Excel's form (spaces ignored), or a number. */
function parse(v: Value): Complex | CellError {
  if (isError(v as Scalar)) return v as CellError
  if (typeof v === "boolean") return VALUE
  if (typeof v === "number") return { re: v, im: 0, suffix: null }
  if (v === null) return { re: 0, im: 0, suffix: null }
  if (typeof v !== "string") return VALUE
  const text = v.replace(/ /g, "")
  let re = 0
  let im = 0
  let suffix: Suffix | null = null
  let m: RegExpExecArray | null
  if ((m = REAL.exec(text))) re = readNumber(m[1])
  else if ((m = IMAG.exec(text))) {
    im = (m[1] === "-" ? -1 : 1) * (m[2] === undefined ? 1 : readNumber(m[2]))
    suffix = m[3] as Suffix
  } else if ((m = BOTH.exec(text))) {
    re = readNumber(m[1])
    const sign = (m[2] === "-" ? -1 : 1) * (m[3] === "-" ? -1 : 1)
    im = sign * (m[4] === undefined ? 1 : readNumber(m[4]))
    suffix = m[5] as Suffix
  } else return NUM
  if (!Number.isFinite(re) || !Number.isFinite(im)) return NUM
  return { re: re + 0, im: im + 0, suffix }
}

function part(x: number): string {
  return numberToText(x, 21)
}

/** Excel's text for a complex number: "3+4i", "i", "-2.5j", "7". */
export function formatComplex(re: number, im: number, suffix: Suffix = "i"): string | CellError {
  if (!Number.isFinite(re) || !Number.isFinite(im)) return NUM
  const imText = part(Math.abs(im))
  const imPart = (imText === "1" ? "" : imText) + suffix
  if (im === 0 || imText === "0") return part(re)
  if (part(re) === "0") return (im < 0 ? "-" : "") + imPart
  return part(re) + (im < 0 ? "-" : "+") + imPart
}

function out(z: { re: number; im: number }, suffix: Suffix | null): string | CellError {
  return formatComplex(z.re, z.im, suffix ?? "i")
}

// ── Sine and cosine as Excel computes them ───────────────────────────────────────────────────────────────────────

/** π/2 in three parts; the last is Excel's (6.1257…e-17 against the true 6.1232…e-17). */
const HALF_PI_1 = 1.5707963267341256
const HALF_PI_2 = 1.5707963267948966 - HALF_PI_1
const HALF_PI_3 = 6.1257422745431e-17
/** Excel's limit for the argument of SIN and COS. */
const TRIG_LIMIT = 2 ** 27

function reduce(x: number): [number, number] {
  const k = Math.round(x / 1.5707963267948966)
  return [x - k * HALF_PI_1 - k * HALF_PI_2 - k * HALF_PI_3, ((k % 4) + 4) % 4]
}
function sin(x: number): number {
  const [r, q] = reduce(x)
  return q === 0 ? Math.sin(r) : q === 1 ? Math.cos(r) : q === 2 ? -Math.sin(r) : -Math.cos(r)
}
function cos(x: number): number {
  const [r, q] = reduce(x)
  return q === 0 ? Math.cos(r) : q === 1 ? -Math.sin(r) : q === 2 ? -Math.cos(r) : Math.sin(r)
}

// ── Arithmetic ───────────────────────────────────────────────────────────────────────────────────────────────────

type Z = { re: number; im: number }
const mul = (a: Z, b: Z): Z => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re })
function div(a: Z, b: Z): Z | CellError {
  const d = b.re * b.re + b.im * b.im
  if (d === 0) return NUM
  return { re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d }
}
const finite = (z: Z): Z | CellError => (Number.isFinite(z.re) && Number.isFinite(z.im) ? z : NUM)

function sinZ(z: Z): Z {
  return { re: sin(z.re) * Math.cosh(z.im), im: cos(z.re) * Math.sinh(z.im) }
}
function cosZ(z: Z): Z {
  return { re: cos(z.re) * Math.cosh(z.im), im: -sin(z.re) * Math.sinh(z.im) }
}
function sinhZ(z: Z): Z {
  return { re: Math.sinh(z.re) * cos(z.im), im: Math.cosh(z.re) * sin(z.im) }
}
function coshZ(z: Z): Z {
  return { re: Math.cosh(z.re) * cos(z.im), im: Math.sinh(z.re) * sin(z.im) }
}
/** tan z = (sin 2a + i sinh 2b) / (cos 2a + cosh 2b), which tends to ±i for a large imaginary part. */
function tanZ(z: Z): Z {
  const b2 = 2 * z.im
  if (Math.abs(b2) > 700) return { re: 0, im: Math.sign(b2) }
  const d = cos(2 * z.re) + Math.cosh(b2)
  return { re: sin(2 * z.re) / d, im: Math.sinh(b2) / d }
}

// ── Registration ─────────────────────────────────────────────────────────────────────────────────────────────────

scalarFn("COMPLEX", 2, 3, (a) => {
  const re = atpNumber(a[0])
  if (isError(re)) return re
  const im = atpNumber(a[1])
  if (isError(im)) return im
  const suffix = a.length > 2 && a[2] !== null && a[2] !== "" ? a[2] : "i"
  if (suffix !== "i" && suffix !== "j") return VALUE
  return formatComplex(re, im, suffix)
})

/** A function of one complex number. `limit` lists the parts that go into a sine or cosine (Excel's 2^27 cap). */
function unary(name: string, f: (z: Complex) => Z | Scalar | CellError, limit: "re" | "im" | "both" | null = null) {
  scalarFn(name, 1, 1, (a) => {
    const z = parse(a[0])
    if (isError(z)) return z
    if (limit && ((limit !== "im" && Math.abs(z.re) >= TRIG_LIMIT) || (limit !== "re" && Math.abs(z.im) >= TRIG_LIMIT)))
      return NUM
    const r = f(z)
    if (r === null || typeof r !== "object" || isError(r as Scalar)) return r as Scalar
    const w = finite(r as Z)
    return isError(w) ? w : out(w, z.suffix)
  })
}

const isZero = (z: Z) => z.re === 0 && z.im === 0

unary("IMREAL", (z) => z.re)
unary("IMAGINARY", (z) => z.im)
unary("IMABS", (z) => Math.hypot(z.re, z.im))
unary("IMARGUMENT", (z) => (isZero(z) ? DIV0 : Math.atan2(z.im, z.re)))
unary("IMCONJUGATE", (z) => ({ re: z.re, im: -z.im }))
unary("IMSIN", sinZ, "re")
unary("IMCOS", cosZ, "re")
unary("IMTAN", tanZ, "re")
unary("IMSINH", sinhZ, "im")
unary("IMCOSH", coshZ, "im")
// sec z = 2(cos a cosh b + i sin a sinh b) / (cos 2a + cosh 2b), which gives Excel's last digit.
unary(
  "IMSEC",
  (z) => {
    if (z.im === 0) return { re: 1 / cos(z.re), im: 0 }
    const d = cos(2 * z.re) + Math.cosh(2 * z.im)
    return { re: (2 * cos(z.re) * Math.cosh(z.im)) / d, im: (2 * sin(z.re) * Math.sinh(z.im)) / d }
  },
  "re",
)
unary("IMCSC", (z) => (isZero(z) ? NUM : div({ re: 1, im: 0 }, sinZ(z))), "re")
unary("IMCOT", (z) => (isZero(z) ? NUM : div({ re: 1, im: 0 }, tanZ(z))), "re")
unary("IMSECH", (z) => div({ re: 1, im: 0 }, coshZ(z)), "im")
unary("IMCSCH", (z) => (isZero(z) ? NUM : div({ re: 1, im: 0 }, sinhZ(z))), "im")
unary(
  "IMEXP",
  (z) => {
    const r = Math.exp(z.re)
    return { re: r * cos(z.im), im: r * sin(z.im) }
  },
  "im",
)
function ln(z: Z): Z | CellError {
  if (isZero(z)) return NUM
  return { re: Math.log(Math.hypot(z.re, z.im)), im: Math.atan2(z.im, z.re) }
}
unary("IMLN", ln)
unary("IMLOG10", (z) => {
  const w = ln(z)
  return isError(w) ? w : { re: w.re / Math.LN10, im: w.im / Math.LN10 }
})
unary("IMLOG2", (z) => {
  const w = ln(z)
  return isError(w) ? w : { re: w.re / Math.LN2, im: w.im / Math.LN2 }
})

/** z^n in polar form, with the modulus as Excel takes it (√(a²+b²), so tiny numbers underflow to 0). */
function power(z: Z, n: number): Z | CellError {
  const r = Math.sqrt(z.re * z.re + z.im * z.im)
  if (r === 0) return n > 0 ? { re: 0, im: 0 } : NUM
  const theta = Math.atan2(z.im, z.re)
  const rn = Math.pow(r, n)
  return { re: rn * cos(n * theta), im: rn * sin(n * theta) }
}
unary("IMSQRT", (z) => {
  const r = Math.sqrt(z.re * z.re + z.im * z.im)
  return r === 0 ? { re: 0, im: 0 } : power(z, 0.5)
})
scalarFn("IMPOWER", 2, 2, (a) => {
  const z = parse(a[0])
  if (isError(z)) return z
  const n = atpNumber(a[1])
  if (isError(n)) return n
  const w = power(z, n)
  if (isError(w)) return w
  const f = finite(w)
  return isError(f) ? f : out(f, z.suffix)
})

/** The suffix two numbers share (a number without an imaginary part takes either); null when they differ. */
function joinSuffix(a: Suffix | null, b: Suffix | null): Suffix | null | false {
  if (a && b && a !== b) return false
  return a ?? b
}

for (const [name, op, mismatch] of [
  ["IMSUB", (x: Z, y: Z): Z | CellError => ({ re: x.re - y.re, im: x.im - y.im }), NUM],
  ["IMDIV", div, NUM],
] as const) {
  scalarFn(name, 2, 2, (a) => {
    const x = parse(a[0])
    if (isError(x)) return x
    const y = parse(a[1])
    if (isError(y)) return y
    const suffix = joinSuffix(x.suffix, y.suffix)
    if (suffix === false) return mismatch
    const w = op(x, y)
    if (isError(w)) return w
    const f = finite(w)
    return isError(f) ? f : out(f, suffix)
  })
}

/** IMSUM and IMPRODUCT: any number of complex numbers, ranges and arrays; blanks are skipped. */
for (const [name, start, op] of [
  ["IMSUM", { re: 0, im: 0 }, (x: Z, y: Z): Z => ({ re: x.re + y.re, im: x.im + y.im })],
  ["IMPRODUCT", { re: 1, im: 0 }, mul],
] as const) {
  define(name, {
    min: 1,
    max: 255,
    call: (args, ctx) => {
      let acc: Z = start
      let suffix: Suffix | null = null
      let error: CellError | null = null
      forEachScalar(args, ctx, (v) => {
        const z = parse(v)
        if (isError(z)) {
          error = z
          return false
        }
        const s = joinSuffix(suffix, z.suffix)
        if (s === false) {
          error = VALUE
          return false
        }
        suffix = s
        acc = op(acc, z)
      })
      if (error) return error
      const f = finite(acc)
      return isError(f) ? f : out(f, suffix)
    },
  })
}
