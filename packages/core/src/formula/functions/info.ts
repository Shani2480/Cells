import { ERROR_CODES, isError } from "../../values"
import { NA, toNumber, VALUE } from "../coerce"
import { toScalar } from "../evaluate"
import { isMatrix, isRef, type Scalar } from "../types"
import { define } from "./registry"

const is = (name: string, test: (v: Scalar) => boolean) =>
  define(name, { min: 1, max: 1, params: ["scalar"], errors: true, call: (a) => test(a[0] as Scalar) })

is("ISNUMBER", (v) => typeof v === "number")
is("ISTEXT", (v) => typeof v === "string")
is("ISNONTEXT", (v) => typeof v !== "string")
is("ISLOGICAL", (v) => typeof v === "boolean")
is("ISERROR", (v) => isError(v))
is("ISERR", (v) => isError(v) && v.error !== "#N/A")
is("ISNA", (v) => isError(v) && v.error === "#N/A")
/** ISEVEN/ISODD take numbers (and number text); other text and logicals are #VALUE!. */
const parity = (name: string, odd: boolean) =>
  define(name, {
    min: 1,
    max: 1,
    params: ["scalar"],
    call: (a) => {
      const v = a[0] as Scalar
      if (typeof v === "boolean") return VALUE
      const n = toNumber(v)
      if (isError(n)) return n
      return Math.abs(Math.trunc(n)) % 2 === (odd ? 1 : 0)
    },
  })
parity("ISEVEN", false)
parity("ISODD", true)

/** ISBLANK is true only for an empty cell (not for "" from a formula). */
define("ISBLANK", {
  min: 1,
  max: 1,
  params: ["scalar"],
  errors: true,
  call: (a) => a[0] === null,
})
define("ISREF", {
  min: 1,
  max: 1,
  errors: true,
  call: (a) => isRef(a[0]),
})
define("ISFORMULA", {
  min: 1,
  max: 1,
  call: (a, ctx) => {
    if (!isRef(a[0])) return VALUE
    const area = a[0].areas[0]
    return ctx.isFormula(area.sheet, area.r0, area.c0)
  },
})
define("FORMULATEXT", {
  min: 1,
  max: 1,
  call: (a, ctx) => {
    if (!isRef(a[0])) return VALUE
    const area = a[0].areas[0]
    return ctx.formulaText(area.sheet, area.r0, area.c0) ?? NA
  },
})
define("N", {
  min: 1,
  max: 1,
  errors: true,
  call: (a, ctx) => {
    const v = toScalar(a[0], ctx)
    if (isError(v)) return v
    if (typeof v === "number") return v
    return v === true ? 1 : 0
  },
})
define("NA", { min: 0, max: 0, call: () => NA })
define("TYPE", {
  min: 1,
  max: 1,
  errors: true,
  call: (a, ctx) => {
    const v = a[0]
    if (
      isMatrix(v) ||
      (isRef(v) && !(v.areas.length === 1 && v.areas[0].r0 === v.areas[0].r1 && v.areas[0].c0 === v.areas[0].c1))
    )
      return 64
    const s = toScalar(v, ctx)
    if (s === null || typeof s === "number") return 1
    if (typeof s === "string") return 2
    if (typeof s === "boolean") return 4
    return 16
  },
})
define("ERROR.TYPE", {
  min: 1,
  max: 1,
  params: ["scalar"],
  errors: true,
  call: (a) => {
    const v = a[0] as Scalar
    if (!isError(v)) return NA
    const codes = [
      "#NULL!",
      "#DIV/0!",
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#NUM!",
      "#N/A",
      "#GETTING_DATA",
      "",
      "",
      "",
      "",
      "",
      "#SPILL!",
      "#CALC!",
    ]
    const i = codes.indexOf(v.error)
    return i >= 0 ? i + 1 : NA
  },
})
define("SHEET", {
  min: 0,
  max: 1,
  call: (a, ctx) => {
    if (a.length === 0) return ctx.sheet + 1
    const v = a[0]
    if (isRef(v)) return v.areas[0].sheet + 1
    if (typeof v === "string") {
      const i = ctx.sheetIndex(v)
      return i >= 0 ? i + 1 : NA
    }
    return VALUE
  },
})
define("SHEETS", {
  min: 0,
  max: 1,
  call: (a, ctx) => {
    if (a.length === 0) return ctx.sheetCount()
    const v = a[0]
    return isRef(v) ? new Set(v.areas.map((x) => x.sheet)).size : VALUE
  },
})

export const KNOWN_ERRORS = ERROR_CODES
