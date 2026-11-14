import { isError, type CellError } from "../../values"
import { compareScalars, NA, toBool, VALUE } from "../coerce"
import { isSingle, toMatrix, toScalar } from "../evaluate"
import { isMatrix, isRef, type Matrix, type Scalar, type Value } from "../types"
import { forEachScalar } from "./helpers"
import { define } from "./registry"

/** A condition, lifted over arrays: IF(A1:A3>1, …) spills one result per element. */
function liftCondition(cond: Value, ctx: Parameters<typeof toScalar>[1], pick: (c: Scalar) => Value): Value {
  if (isSingle(cond)) return pick(toScalar(cond, ctx))
  const m = toMatrix(cond, ctx)
  if (isError(m)) return m
  // Branches can be arrays too; each element takes the matching element of the chosen branch.
  const results = m.data.map((c) => pick(c))
  const parts: Matrix[] = []
  for (const r of results) {
    const rm = toMatrix(r, ctx)
    parts.push(isError(rm) ? { kind: "matrix", rows: 1, cols: 1, data: [rm] } : rm)
  }
  return {
    kind: "matrix",
    rows: m.rows,
    cols: m.cols,
    data: m.data.map((_, i) => {
      const part = parts[i]
      const r = Math.floor(i / m.cols)
      const c = i % m.cols
      const rr = part.rows === 1 ? 0 : r
      const cc = part.cols === 1 ? 0 : c
      return rr < part.rows && cc < part.cols ? part.data[rr * part.cols + cc] : NA
    }),
  }
}

define("IF", {
  min: 1,
  max: 3,
  lazy: (args, ctx) => {
    const cond = args[0]()
    return liftCondition(cond, ctx, (c) => {
      if (isError(c)) return c
      const b = toBool(c)
      if (isError(b)) return b
      if (b) return args.length > 1 ? (args[1]() ?? 0) : true
      return args.length > 2 ? (args[2]() ?? 0) : false
    })
  },
})

define("IFS", {
  min: 2,
  max: 254,
  lazy: (args, ctx) => {
    if (args.length % 2 !== 0) return VALUE
    for (let i = 0; i < args.length; i += 2) {
      const c = toScalar(args[i](), ctx)
      if (isError(c)) return c
      const b = toBool(c)
      if (isError(b)) return b
      if (b) return args[i + 1]() ?? 0
    }
    return NA
  },
})

function errorFallback(onlyNA: boolean) {
  return (args: (() => Value)[], ctx: Parameters<typeof toScalar>[1]): Value => {
    const value = args[0]()
    const test = (v: Scalar) => isError(v) && (!onlyNA || v.error === "#N/A")
    if (isSingle(value)) {
      const v = toScalar(value, ctx)
      return test(v) ? (args[1]() ?? 0) : isRef(value) ? value : v === null ? 0 : v
    }
    const m = toMatrix(value, ctx)
    if (isError(m)) return test(m) ? args[1]() : m
    let fallback: Scalar | undefined
    return {
      kind: "matrix",
      rows: m.rows,
      cols: m.cols,
      data: m.data.map((v) => {
        if (!test(v)) return v
        fallback ??= toScalar(args[1](), ctx)
        return fallback
      }),
    }
  }
}
define("IFERROR", { min: 2, max: 2, lazy: errorFallback(false) })
define("IFNA", { min: 2, max: 2, lazy: errorFallback(true) })

/** AND/OR/XOR over logicals: text in cells is ignored, typed text that is not TRUE/FALSE is #VALUE!. */
function logicalReduce(
  args: Value[],
  ctx: Parameters<typeof forEachScalar>[1],
  step: (acc: boolean, b: boolean) => boolean,
  init: boolean,
): Scalar {
  let acc = init
  let seen = false
  let error: CellError | null = null
  forEachScalar(args, ctx, (v, direct) => {
    if (isError(v)) {
      error = v
      return false
    }
    if (typeof v === "string" && !direct) return
    const b = toBool(v)
    if (isError(b)) {
      // Text typed as an argument counts only when it reads as TRUE/FALSE; other text is skipped like text in
      // a range (AND(TRUE,"x") is TRUE), and only text alone gives #VALUE!.
      if (typeof v === "string") return
      error = b
      return false
    }
    seen = true
    acc = step(acc, b)
  })
  if (error) return error
  return seen ? acc : VALUE
}
define("AND", { min: 1, max: 255, call: (a, ctx) => logicalReduce(a, ctx, (x, b) => x && b, true) })
define("OR", { min: 1, max: 255, call: (a, ctx) => logicalReduce(a, ctx, (x, b) => x || b, false) })
define("XOR", { min: 1, max: 255, call: (a, ctx) => logicalReduce(a, ctx, (x, b) => x !== b, false) })
define("NOT", {
  min: 1,
  max: 1,
  params: ["scalar"],
  call: (a) => {
    const b = toBool(a[0] as Scalar)
    return isError(b) ? b : !b
  },
})
define("TRUE", { min: 0, max: 0, call: () => true })
define("FALSE", { min: 0, max: 0, call: () => false })

define("SWITCH", {
  min: 3,
  max: 254,
  lazy: (args, ctx) => {
    const v = toScalar(args[0](), ctx)
    if (isError(v)) return v
    let i = 1
    for (; i + 1 < args.length; i += 2) {
      const candidate = toScalar(args[i](), ctx)
      if (isError(candidate)) return candidate
      if (v !== null && candidate !== null && compareScalars(v, candidate) === 0 && typeof v === typeof candidate) {
        return args[i + 1]() ?? 0
      }
    }
    return i < args.length ? (args[i]() ?? 0) : NA
  },
})

define("CHOOSE", {
  min: 2,
  max: 255,
  lazy: (args, ctx) => {
    const index = args[0]()
    return liftCondition(index, ctx, (v) => {
      if (isError(v)) return v
      const n = typeof v === "number" ? Math.trunc(v) : typeof v === "boolean" ? Number(v) : NaN
      if (typeof v === "string") {
        const t = Number(v)
        if (!Number.isFinite(t)) return VALUE
        return pickChoice(args, Math.trunc(t))
      }
      return pickChoice(args, n)
    })
  },
})
function pickChoice(args: (() => Value)[], n: number): Value {
  if (!(n >= 1 && n < args.length)) return VALUE
  return args[n]() ?? 0
}

export function asMatrixOrError(v: Value): v is Matrix {
  return isMatrix(v)
}
