/**
 * Engineering functions: number bases (BIN2DEC … HEX2OCT), BIT*, ERF/ERFC, DELTA, GESTEP, the Bessel functions and
 * CONVERT. Complex numbers are in complex.ts.
 *
 * These came from Excel's Analysis ToolPak and keep its habits: logicals are #VALUE!, an omitted argument is #N/A,
 * and the Bessel functions use the classic rational approximations (Numerical Recipes), good to about 1e-8, which is
 * what Excel returns (BESSELJ(0,0) is 1.00000000283141 there too).
 */
import { isError, type CellError } from "../../values"
import { NA, NUM, numberToText, toNumber, VALUE } from "../coerce"
import type { Scalar, Value } from "../types"
import { num } from "./helpers"
import { scalarFn } from "./registry"
import { erf, erfc } from "./special"

/** A number argument of an Analysis ToolPak function: text that reads as a number is fine, a logical is #VALUE!. */
export function atpNumber(v: Value): number | CellError {
  if (typeof v === "boolean") return VALUE
  return toNumber(v as Scalar)
}

/** Denormal results show as 0 in Excel. */
function flushed(x: number): number | CellError {
  if (!Number.isFinite(x)) return NUM
  return Math.abs(x) < 2.2250738585072014e-308 ? 0 : x
}

// ── Number bases ─────────────────────────────────────────────────────────────────────────────────────────────────

type Base = { radix: number; digits: RegExp; min: number; max: number }
const BASES: Record<"BIN" | "OCT" | "HEX", Base> = {
  BIN: { radix: 2, digits: /^[01]*$/, min: -512, max: 511 },
  OCT: { radix: 8, digits: /^[0-7]*$/, min: -(2 ** 29), max: 2 ** 29 - 1 },
  HEX: { radix: 16, digits: /^[0-9A-Fa-f]*$/, min: -(2 ** 39), max: 2 ** 39 - 1 },
}

/** Ten digits at most; ten digits with the top bit set are a negative number in two's complement. */
function parseBase(v: Value, base: Base): number | CellError {
  if (typeof v === "boolean") return VALUE
  if (isError(v as Scalar)) return v as CellError
  const text = v === null ? "" : typeof v === "number" ? numberToText(v) : String(v)
  if (text.length > 10 || !base.digits.test(text)) return NUM
  if (text === "") return 0
  const n = parseInt(text, base.radix)
  return text.length === 10 && n >= base.radix ** 10 / 2 ? n - base.radix ** 10 : n
}

function formatBase(n: number, base: Base, places: Value | undefined): string | CellError {
  if (n < base.min || n > base.max) return NUM
  if (n < 0) return (base.radix ** 10 + n).toString(base.radix).toUpperCase()
  const text = n.toString(base.radix).toUpperCase()
  if (places === undefined) return text
  const p = atpNumber(places)
  if (isError(p)) return p
  const width = Math.trunc(p)
  if (width < text.length || width > 10) return NUM
  return text.padStart(width, "0")
}

for (const from of ["BIN", "OCT", "HEX", "DEC"] as const) {
  for (const to of ["BIN", "OCT", "HEX", "DEC"] as const) {
    if (from === to) continue
    const name = `${from}2${to}`
    if (to === "DEC") {
      scalarFn(name, 1, 1, (a) => parseBase(a[0], BASES[from as "BIN"]), { omittedIsNA: true })
      continue
    }
    scalarFn(
      name,
      1,
      2,
      (a) => {
        if (isError(a[0] as Scalar)) return a[0] as CellError
        let n: number | CellError
        if (from === "DEC") {
          const x = atpNumber(a[0])
          n = isError(x) ? x : Math.trunc(x)
        } else n = parseBase(a[0], BASES[from])
        if (isError(n)) return n
        return formatBase(n, BASES[to], a.length > 1 ? a[1] : undefined)
      },
      { omittedIsNA: true },
    )
  }
}

// ── BIT functions ────────────────────────────────────────────────────────────────────────────────────────────────

const BIT_LIMIT = 2 ** 48

function bitArg(v: Value): bigint | CellError {
  const n = num(v)
  if (isError(n)) return n
  if (!Number.isInteger(n) || n < 0 || n >= BIT_LIMIT) return NUM
  return BigInt(n)
}

for (const [name, op] of [
  ["BITAND", (x: bigint, y: bigint) => x & y],
  ["BITOR", (x: bigint, y: bigint) => x | y],
  ["BITXOR", (x: bigint, y: bigint) => x ^ y],
] as const) {
  scalarFn(name, 2, 2, (a) => {
    const x = bitArg(a[0])
    if (isError(x)) return x
    const y = bitArg(a[1])
    if (isError(y)) return y
    return Number(op(x, y))
  })
}

function shift(v: Value, amount: Value, left: boolean): number | CellError {
  const n = bitArg(v)
  if (isError(n)) return n
  const s = num(amount)
  if (isError(s)) return s
  let by = Math.trunc(s)
  if (Math.abs(by) > 53) return NUM
  if (!left) by = -by
  if (by < 0) return Number(n >> BigInt(-by))
  const result = n << BigInt(by)
  // Excel keeps 53 bits and then checks the 48-bit limit (BITLSHIFT(5,53) is 0, BITLSHIFT(1,48) is #NUM!).
  const kept = result & ((1n << 53n) - 1n)
  return kept >= BigInt(BIT_LIMIT) ? NUM : Number(kept)
}
scalarFn("BITLSHIFT", 2, 2, (a) => shift(a[0], a[1], true))
scalarFn("BITRSHIFT", 2, 2, (a) => shift(a[0], a[1], false))

// ── ERF, ERFC, DELTA, GESTEP ─────────────────────────────────────────────────────────────────────────────────────

scalarFn(
  "ERF",
  1,
  2,
  (a) => {
    const lower = atpNumber(a[0])
    if (isError(lower)) return lower
    if (a.length < 2) return erf(lower)
    const upper = atpNumber(a[1])
    if (isError(upper)) return upper
    return erf(upper) - erf(lower)
  },
  { omittedIsNA: true },
)
scalarFn("ERF.PRECISE", 1, 1, (a) => {
  const x = atpNumber(a[0])
  return isError(x) ? x : erf(x)
})
for (const name of ["ERFC", "ERFC.PRECISE"])
  scalarFn(
    name,
    1,
    1,
    (a) => {
      const x = atpNumber(a[0])
      return isError(x) ? x : flushed(erfc(x))
    },
    { omittedIsNA: name === "ERFC" },
  )

scalarFn(
  "DELTA",
  1,
  2,
  (a) => {
    const x = atpNumber(a[0])
    if (isError(x)) return x
    const y = a.length > 1 ? atpNumber(a[1]) : 0
    if (isError(y)) return y
    return x === y ? 1 : 0
  },
  { omittedIsNA: true },
)
scalarFn(
  "GESTEP",
  1,
  2,
  (a) => {
    const x = atpNumber(a[0])
    if (isError(x)) return x
    const step = a.length > 1 ? atpNumber(a[1]) : 0
    if (isError(step)) return step
    return x >= step ? 1 : 0
  },
  { omittedIsNA: true },
)

// ── Bessel functions (Numerical Recipes in C, 2nd edition, §6.5 and §6.6, as Excel computes them) ───────────────

function besselJ0(x: number): number {
  const ax = Math.abs(x)
  if (ax < 8) {
    const y = x * x
    const p =
      57568490574.0 +
      y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * -184.9052456))))
    const q = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y * 1.0))))
    return p / q
  }
  const z = 8 / ax
  const y = z * z
  const xx = ax - 0.785398164
  const p = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
  const q =
    -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934935152e-7)))
  return Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * p - z * Math.sin(xx) * q)
}

function besselJ1(x: number): number {
  const ax = Math.abs(x)
  if (ax < 8) {
    const y = x * x
    const p =
      x *
      (72362614232.0 +
        y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.4826 + y * -30.16036606)))))
    const q = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y * 1.0))))
    return p / q
  }
  const z = 8 / ax
  const y = z * z
  const xx = ax - 2.356194491
  const p = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)))
  const q = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
  const ans = Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * p - z * Math.sin(xx) * q)
  return x < 0 ? -ans : ans
}

const BIGNO = 1.0e10
const BIGNI = 1.0e-10

function besselJ(n: number, x: number): number {
  if (n === 0) return besselJ0(x)
  if (n === 1) return besselJ1(x)
  const ax = Math.abs(x)
  if (ax === 0) return 0
  const tox = 2 / ax
  let ans: number
  if (ax > n) {
    // Upward recurrence from J0 and J1.
    let bjm = besselJ0(ax)
    let bj = besselJ1(ax)
    for (let j = 1; j < n; j++) {
      const bjp = j * tox * bj - bjm
      bjm = bj
      bj = bjp
    }
    ans = bj
  } else {
    // Miller's downward recurrence, normalised with J0 + 2(J2 + J4 + …) = 1.
    const m = 2 * Math.trunc((n + Math.trunc(Math.sqrt(40 * n))) / 2)
    let even = false
    let bjp = 0
    let sum = 0
    let bj = 1
    ans = 0
    for (let j = m; j > 0; j--) {
      const bjm = j * tox * bj - bjp
      bjp = bj
      bj = bjm
      if (Math.abs(bj) > BIGNO) {
        bj *= BIGNI
        bjp *= BIGNI
        ans *= BIGNI
        sum *= BIGNI
      }
      if (even) sum += bj
      even = !even
      if (j === n) ans = bjp
    }
    sum = 2 * sum - bj
    ans /= sum
  }
  return x < 0 && n % 2 === 1 ? -ans : ans
}

function besselY0(x: number): number {
  if (x < 8) {
    const y = x * x
    const p =
      -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))))
    const q = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.2647 + y * (226.1030244 + y * 1.0))))
    return p / q + 0.636619772 * besselJ0(x) * Math.log(x)
  }
  const z = 8 / x
  const y = z * z
  const xx = x - 0.785398164
  const p = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
  const q =
    -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * -0.934945152e-7)))
  return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * p + z * Math.cos(xx) * q)
}

function besselY1(x: number): number {
  if (x < 8) {
    const y = x * x
    const p =
      x *
      (-0.4900604943e13 +
        y *
          (0.127527439e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))))
    const q =
      0.249958057e14 +
      y *
        (0.4244419664e12 +
          y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.102042605e6 + y * (0.3549632885e3 + y)))))
    return p / q + 0.636619772 * (besselJ1(x) * Math.log(x) - 1 / x)
  }
  const z = 8 / x
  const y = z * z
  const xx = x - 2.356194491
  const p = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)))
  const q = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
  return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * p + z * Math.cos(xx) * q)
}

function besselY(n: number, x: number): number {
  if (n === 0) return besselY0(x)
  if (n === 1) return besselY1(x)
  const tox = 2 / x
  let by = besselY1(x)
  let bym = besselY0(x)
  for (let j = 1; j < n; j++) {
    const byp = j * tox * by - bym
    bym = by
    by = byp
  }
  return by
}

function besselI0(x: number): number {
  const ax = Math.abs(x)
  if (ax < 3.75) {
    let y = x / 3.75
    y *= y
    return (
      1.0 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))))
    )
  }
  const y = 3.75 / ax
  return (
    (Math.exp(ax) / Math.sqrt(ax)) *
    (0.39894228 +
      y *
        (0.1328592e-1 +
          y *
            (0.225319e-2 +
              y *
                (-0.157565e-2 +
                  y *
                    (0.916281e-2 + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))))
  )
}

function besselI1(x: number): number {
  const ax = Math.abs(x)
  let ans: number
  if (ax < 3.75) {
    let y = x / 3.75
    y *= y
    ans =
      ax *
      (0.5 +
        y *
          (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))))
  } else {
    const y = 3.75 / ax
    ans = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2))
    ans = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * ans))))
    ans *= Math.exp(ax) / Math.sqrt(ax)
  }
  return x < 0 ? -ans : ans
}

function besselI(n: number, x: number): number {
  if (n === 0) return besselI0(x)
  if (n === 1) return besselI1(x)
  if (x === 0) return 0
  const tox = 2 / Math.abs(x)
  let bip = 0
  let bi = 1
  let ans = 0
  for (let j = 2 * (n + Math.trunc(Math.sqrt(40 * n))); j > 0; j--) {
    const bim = bip + j * tox * bi
    bip = bi
    bi = bim
    if (Math.abs(bi) > BIGNO) {
      ans *= BIGNI
      bi *= BIGNI
      bip *= BIGNI
    }
    if (j === n) ans = bip
  }
  ans *= besselI0(x) / bi
  return x < 0 && n % 2 === 1 ? -ans : ans
}

function besselK0(x: number): number {
  if (x <= 2) {
    const y = (x * x) / 4
    return (
      -Math.log(x / 2) * besselI0(x) +
      (-0.57721566 +
        y * (0.4227842 + y * (0.23069756 + y * (0.348859e-1 + y * (0.262698e-2 + y * (0.1075e-3 + y * 0.74e-5))))))
    )
  }
  const y = 2 / x
  return (
    (Math.exp(-x) / Math.sqrt(x)) *
    (1.25331414 +
      y *
        (-0.7832358e-1 +
          y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.25154e-2 + y * 0.53208e-3))))))
  )
}

function besselK1(x: number): number {
  if (x <= 2) {
    const y = (x * x) / 4
    return (
      Math.log(x / 2) * besselI1(x) +
      (1 / x) *
        (1.0 +
          y *
            (0.15443144 +
              y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4))))))
    )
  }
  const y = 2 / x
  return (
    (Math.exp(-x) / Math.sqrt(x)) *
    (1.25331414 +
      y *
        (0.23498619 +
          y * (-0.365562e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))))
  )
}

function besselK(n: number, x: number): number {
  if (n === 0) return besselK0(x)
  if (n === 1) return besselK1(x)
  const tox = 2 / x
  let bkm = besselK0(x)
  let bk = besselK1(x)
  for (let j = 1; j < n; j++) {
    const bkp = bkm + j * tox * bk
    bkm = bk
    bk = bkp
  }
  return bk
}

for (const [name, fn, positive] of [
  ["BESSELJ", besselJ, false],
  ["BESSELY", besselY, true],
  ["BESSELI", besselI, false],
  ["BESSELK", besselK, true],
] as const) {
  scalarFn(
    name,
    2,
    2,
    (a) => {
      const x = atpNumber(a[0])
      if (isError(x)) return x
      const order = atpNumber(a[1])
      if (isError(order)) return order
      const n = Math.trunc(order)
      if (n < 0 || (positive && x <= 0)) return NUM
      const v = fn(n, x)
      return Number.isFinite(v) ? v : NUM
    },
    { omittedIsNA: true },
  )
}

// ── CONVERT ──────────────────────────────────────────────────────────────────────────────────────────────────────

type Unit = { category: string; factor: number; prefixes: boolean; power: number }

/**
 * Each unit's size in its category's reference unit (g, m, sec, Pa, N, J, W, T, l, m², bit, m/s), as Excel
 * reports them (CONVERT(1, unit, reference)); temperatures are worked out separately. "+" marks units that take a
 * prefix; 2 and 3 after the name are its power, so a prefix applies squared or cubed.
 */
const UNIT_TABLE: Record<string, string> = {
  mass: "g+ 1|sg 14593.902937206363|lbm 453.59237|u+ 1.660538782e-24|ozm 28.349523125|grain 0.06479891|cwt 45359.237|shweight 45359.237|uk_cwt 50802.34544|lcwt 50802.34544|hweight 50802.34544|stone 6350.29318|ton 907184.74|uk_ton 1016046.9088|LTON 1016046.9088|brton 1016046.9088",
  distance:
    "m+ 1|mi 1609.344|Nmi 1852|in 0.0254|ft 0.3048|yd 0.9144|ang+ 1e-10|ell 1.143|ly+ 9460730472580800|parsec+ 3.0856775812815532e16|pc+ 3.0856775812815532e16|Picapt 0.0003527777777777778|Pica 0.0003527777777777778|pica 0.004233333333333334|survey_mi 1609.3472186944373",
  time: "yr 31557600|day 86400|d 86400|hr 3600|mn 60|min 60|sec+ 1|s+ 1",
  pressure: "Pa+ 1|p+ 1|atm+ 101325|at+ 101325|mmHg+ 133.322|psi 6894.757293168362|Torr 133.32236842105263",
  force: "N+ 1|dyn+ 1e-5|dy+ 1e-5|lbf 4.4482216152605005|pond+ 0.00980665",
  energy:
    "J+ 1|e+ 1e-7|c+ 4.184|cal+ 4.1868|eV+ 1.602176487e-19|ev+ 1.602176487e-19|HPh 2684519.537696173|hh 2684519.537696173|Wh+ 3600|wh+ 3600|flb 1.3558179483314003|BTU 1055.05585262|btu 1055.05585262",
  power: "HP 745.6998715822702|h 745.6998715822702|PS 735.49875|W+ 1|w+ 1",
  magnetism: "T+ 1|ga+ 0.0001",
  volume:
    "tsp 0.00492892159375|tspm 0.005|tbs 0.01478676478125|oz 0.0295735295625|cup 0.2365882365|pt 0.473176473|us_pt 0.473176473|uk_pt+ 0.56826125|qt 0.946352946|uk_qt 1.1365225|gal 3.785411784|uk_gal 4.54609|l+ 1|L+ 1|lt+ 1|ang3+3 1e-27|ang^3+3 1e-27|barrel 158.987294928|bushel 35.23907016688|ft3 28.316846592|ft^3 28.316846592|in3 0.016387064|in^3 0.016387064|ly3 8.467866646237152e50|ly^3 8.467866646237152e50|m3+3 1000|m^3+3 1000|mi3 4168181825440.5796|mi^3 4168181825440.5796|yd3 764.554857984|yd^3 764.554857984|Nmi3 6352182208000|Nmi^3 6352182208000|Pica3 4.39039566186557e-8|Pica^3 4.39039566186557e-8|Picapt3 4.39039566186557e-8|Picapt^3 4.39039566186557e-8|GRT 2831.6846592|regton 2831.6846592|MTON 1132.67386368",
  area: "uk_acre 4046.8564224|us_acre 4046.8726098742522|ang2+2 1e-20|ang^2+2 1e-20|ar+ 100|ft2 0.09290304|ft^2 0.09290304|ha 10000|in2 0.00064516|in^2 0.00064516|ly2 8.95054210748189e31|ly^2 8.95054210748189e31|m2+2 1|m^2+2 1|Morgen 2500|mi2 2589988.110336|mi^2 2589988.110336|Nmi2 3429904|Nmi^2 3429904|Picapt2 1.2445216049382715e-7|Pica2 1.2445216049382715e-7|Pica^2 1.2445216049382715e-7|Picapt^2 1.2445216049382715e-7|yd2 0.83612736|yd^2 0.83612736",
  information: "bit+ 1|byte+ 8",
  speed:
    "admkn 0.5147733333333333|kn 0.5144444444444445|m/h+ 0.0002777777777777778|m/hr+ 0.0002777777777777778|m/s+ 1|m/sec+ 1|mph+ 0.44704",
}

const UNITS = new Map<string, Unit>()
for (const [category, list] of Object.entries(UNIT_TABLE)) {
  for (const entry of list.split("|")) {
    const [head, factor] = entry.split(" ")
    const match = /^(.*?)(\+([23])?)?$/.exec(head)!
    UNITS.set(match[1], {
      category,
      factor: Number(factor),
      prefixes: match[2] !== undefined,
      power: match[3] ? Number(match[3]) : 1,
    })
  }
}
const TEMPERATURES = ["C", "cel", "F", "fah", "K", "kel", "Rank", "Reau"]

/** Decimal prefixes as powers of ten ("e" is Excel's other spelling of deka), binary ones as powers of two. */
const PREFIXES: Record<string, number> = {
  Y: 24, Z: 21, E: 18, P: 15, T: 12, G: 9, M: 6, k: 3, h: 2, da: 1, e: 1, d: -1, c: -2, m: -3, u: -6, n: -9, p: -12,
  f: -15, a: -18, z: -21, y: -24,
} // prettier-ignore
const BINARY: Record<string, number> = { ki: 10, Mi: 20, Gi: 30, Ti: 40, Pi: 50, Ei: 60, Zi: 70, Yi: 80 }

type Resolved = { unit: Unit; ten: number; two: number } | { temperature: string; ten: number }

function resolveUnit(name: string): Resolved | null {
  if (TEMPERATURES.includes(name)) return { temperature: name, ten: 0 }
  const exact = UNITS.get(name)
  if (exact) return { unit: exact, ten: 0, two: 0 }
  // A prefix and a unit that takes one: two-letter prefixes first ("dam" is decametre).
  for (const length of [2, 1]) {
    const prefix = name.slice(0, length)
    const rest = name.slice(length)
    if (length === 2 && prefix in BINARY) {
      const unit = UNITS.get(rest)
      if (unit?.category === "information") return { unit, ten: 0, two: BINARY[prefix] }
    }
    if (prefix in PREFIXES) {
      // Kelvin is the one temperature that takes a prefix.
      if (rest === "K" || rest === "kel") return { temperature: "K", ten: PREFIXES[prefix] }
      const unit = UNITS.get(rest)
      if (unit?.prefixes) return { unit, ten: PREFIXES[prefix] * unit.power, two: 0 }
    }
  }
  return null
}

function toKelvin(x: number, unit: string): number {
  switch (unit) {
    case "C":
    case "cel":
      return x + 273.15
    case "F":
    case "fah":
      return ((x - 32) * 5) / 9 + 273.15
    case "Rank":
      return (x * 5) / 9
    case "Reau":
      return x * 1.25 + 273.15
    default:
      return x
  }
}

function convertTemperature(x: number, from: string, to: string): number {
  const norm = (u: string) => ({ cel: "C", fah: "F", kel: "K" })[u] ?? u
  const f = norm(from)
  const t = norm(to)
  if (f === t) return x
  // Between Celsius-based scales directly, so 68 °F is exactly 20 °C.
  if (f === "F" && t === "C") return ((x - 32) * 5) / 9
  if (f === "C" && t === "F") return (x * 9) / 5 + 32
  if (f === "F" && t === "Rank") return x + 459.67
  if (f === "Rank" && t === "F") return x - 459.67
  const k = toKelvin(x, f)
  switch (t) {
    case "C":
      return k - 273.15
    case "F":
      return (k - 273.15) * 1.8 + 32
    case "Rank":
      return k * 1.8
    case "Reau":
      return (k - 273.15) * 0.8
    default:
      return k
  }
}

scalarFn(
  "CONVERT",
  3,
  3,
  (a) => {
    const x = atpNumber(a[0])
    if (isError(x)) return x
    if (typeof a[1] !== "string" || typeof a[2] !== "string") return VALUE_OR_NA(a[1], a[2])
    const from = resolveUnit(a[1])
    const to = resolveUnit(a[2])
    if (!from || !to) return NA
    if ("temperature" in from || "temperature" in to) {
      if (!("temperature" in from && "temperature" in to)) return NA
      const kelvin = from.ten !== 0 || to.ten !== 0
      if (kelvin && (from.temperature !== "K" || to.temperature !== "K")) {
        const k = convertTemperature(x, from.temperature, "K") * 10 ** from.ten
        return convertTemperature(k / 10 ** to.ten, "K", to.temperature)
      }
      return convertTemperature(x, from.temperature, to.temperature) * 10 ** (from.ten - to.ten)
    }
    if (from.unit.category !== to.unit.category) return NA
    let result = (x * from.unit.factor) / to.unit.factor
    if (from.ten !== to.ten) result *= 10 ** (from.ten - to.ten)
    if (from.two !== to.two) result *= 2 ** (from.two - to.two)
    return flushed(result)
  },
  { omittedIsNA: true },
)

/** A unit that isn't text: a logical is #VALUE!, anything else no unit at all. */
function VALUE_OR_NA(a: Value, b: Value): CellError {
  return typeof a === "boolean" || typeof b === "boolean" ? VALUE : NA
}
