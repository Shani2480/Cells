import { isError, type CellError } from "../../values"
import { toNumber, VALUE } from "../coerce"
import type { EvalContext } from "../context"
import { toMatrix } from "../evaluate"
import { isMatrix, isRef, type Area, type Matrix, type Scalar, type Value } from "../types"

/**
 * Visits every scalar of the arguments. `direct` is true for values typed as arguments (SUM(“1”, TRUE) counts
 * them), false for cells of references and elements of arrays (where SUM ignores text and logicals). Blank
 * cells are skipped.
 */
export function forEachScalar(
  args: Value[],
  ctx: EvalContext,
  visit: (value: Scalar, direct: boolean) => void | boolean,
) {
  for (const arg of args) {
    if (isRef(arg)) {
      for (const area of arg.areas) {
        let stop = false
        ctx.values(area, (v) => {
          if (!stop && visit(v, false) === false) stop = true
        })
        if (stop) return
      }
    } else if (isMatrix(arg)) {
      for (const v of arg.data) {
        if (v === null) continue
        if (visit(v, false) === false) return
      }
    } else if (arg !== null) {
      if (visit(arg, true) === false) return
    }
  }
}

/**
 * The numbers of SUM-style arguments: typed text and logicals convert (text that is not a number is #VALUE!),
 * text and logicals in cells and arrays are ignored, the first error wins.
 */
export function collectNumbers(
  args: Value[],
  ctx: EvalContext,
  options: { logicalsInRefs?: boolean } = {},
): number[] | CellError {
  const out: number[] = []
  let error: CellError | null = null
  forEachScalar(args, ctx, (v, direct) => {
    if (isError(v)) {
      error = v
      return false
    }
    if (typeof v === "number") {
      out.push(v)
    } else if (direct) {
      const n = toNumber(v)
      if (isError(n)) {
        error = n
        return false
      }
      out.push(n)
    } else if (options.logicalsInRefs) {
      // AVERAGEA, MAXA…: TRUE is 1, FALSE and text are 0.
      out.push(v === true ? 1 : 0)
    }
  })
  return error ?? out
}

/** A numeric parameter. */
export function num(v: Value): number | CellError {
  return toNumber(v as Scalar)
}

/** An integer parameter (truncated toward zero, as Excel does for counts and positions). */
export function int(v: Value): number | CellError {
  const n = toNumber(v as Scalar)
  return isError(n) ? n : Math.trunc(n)
}

/** A value as one array (a single area, an array constant, or a scalar as 1×1). */
export function arrayArg(v: Value, ctx: EvalContext): Matrix | CellError {
  return toMatrix(v, ctx)
}

export function areaOf(v: Value): Area | null {
  return isRef(v) && v.areas.length === 1 ? v.areas[0] : null
}

export function firstError(...values: (number | CellError | string | boolean)[]): CellError | null {
  for (const v of values) if (isError(v)) return v as CellError
  return null
}

export const notError = <T>(v: T | CellError): v is T => !isError(v)

export { VALUE }
