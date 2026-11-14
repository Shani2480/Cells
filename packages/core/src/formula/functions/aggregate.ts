import { isError } from "../../values"
import { DIV0, VALUE } from "../coerce"
import type { EvalContext } from "../context"
import { isRef, type Scalar, type Value } from "../types"
import { int } from "./helpers"
import { excelSum, SUBTOTAL_SKIP } from "./math"
import { median } from "./stats"
import { FUNCTIONS, define } from "./registry"

/**
 * SUBTOTAL(function_num, refs…): 1–11 include hidden rows, 101–111 skip them. Cells that hold SUBTOTAL or
 * AGGREGATE formulas are skipped so nested subtotals are not counted twice. Hidden rows come from the context.
 */
const NAMES = ["AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "PRODUCT", "STDEV.S", "STDEV.P", "SUM", "VAR.S", "VAR.P"]

/** The values of the reference arguments with nested subtotals (and, when asked, hidden rows) taken out. */
function filteredValues(args: Value[], ctx: EvalContext, skipHidden: boolean, skipErrors: boolean): Value[] | Scalar {
  const out: Value[] = []
  for (const arg of args) {
    if (!isRef(arg)) {
      out.push(arg)
      continue
    }
    const kept: Scalar[] = []
    for (const area of arg.areas) {
      let error: Scalar = null
      ctx.values(area, (v, row, col) => {
        if (error) return
        if (skipHidden && ctx.isHiddenRow?.(area.sheet, row)) return
        const f = ctx.formulaText(area.sheet, row, col)
        if (f && SUBTOTAL_SKIP.test(f)) return
        if (isError(v)) {
          if (skipErrors) return
          error = v
          return
        }
        kept.push(v)
      })
      if (error) return error
    }
    out.push({ kind: "matrix", rows: kept.length || 1, cols: 1, data: kept.length ? kept : [null] })
  }
  return out
}

define("SUBTOTAL", {
  min: 2,
  max: 255,
  call: (a, ctx) => {
    const n = int(a[0])
    if (isError(n)) return n
    const base = n > 100 ? n - 100 : n
    if (base < 1 || base > 11) return VALUE
    const values = filteredValues(a.slice(1), ctx, n > 100, false)
    if (!Array.isArray(values)) return values
    if (values.some((v) => !isRef(v) && !(typeof v === "object" && v !== null && "kind" in v))) return VALUE
    return FUNCTIONS.get(NAMES[base - 1])!.call!(values, ctx)
  },
})

/** AGGREGATE(function, options, ref…) or AGGREGATE(function, options, array, k) for LARGE, SMALL, PERCENTILE… */
const AGG = [
  ...NAMES,
  "MEDIAN",
  "MODE.SNGL",
  "LARGE",
  "SMALL",
  "PERCENTILE.INC",
  "QUARTILE.INC",
  "PERCENTILE.EXC",
  "QUARTILE.EXC",
]
define("AGGREGATE", {
  min: 3,
  max: 253,
  call: (a, ctx) => {
    const fn = int(a[0])
    const options = int(a[1])
    if (isError(fn)) return fn
    if (isError(options)) return options
    if (fn < 1 || fn > 19 || options < 0 || options > 7) return VALUE
    const skipHidden = options === 1 || options === 3 || options === 5 || options === 7
    const skipErrors = options === 2 || options === 3 || options === 6 || options === 7
    const name = AGG[fn - 1]
    const refs = fn >= 14 ? a.slice(2, 3) : a.slice(2)
    const values = filteredValues(refs, ctx, skipHidden, skipErrors)
    if (!Array.isArray(values)) return values
    const extra = fn >= 14 ? a.slice(3, 4) : []
    if (fn >= 14 && extra.length === 0) return VALUE
    return FUNCTIONS.get(name)!.call!([...values, ...extra], ctx)
  },
})

export { DIV0, excelSum, median }
