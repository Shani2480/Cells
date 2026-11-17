import { isError, type CellError } from "../../values"
import { NA, NUM, VALUE } from "../coerce"
import type { Scalar, Value } from "../types"
import type { EvalContext } from "../context"
import {
  detectSeasonality,
  etsConfidence,
  etsForecast,
  etsPosition,
  etsSeries,
  etsStatistic,
  fitEts,
  type EtsSeries,
} from "../../forecast"
import { arrayArg, int, num } from "./helpers"
import { define } from "./registry"

const problem = (p: "num" | "na" | "value"): CellError => (p === "na" ? NA : p === "value" ? VALUE : NUM)

/** An optional integer argument with its default (an omitted or blank argument takes the default). */
function option(v: Value | undefined, fallback: number, min: number, max: number): number | CellError {
  if (v === undefined || v === null) return fallback
  const n = int(v)
  if (isError(n)) return n
  return n < min || n > max ? NUM : n
}

/** The series from (values, timeline) with data completion and aggregation; a CellError when it can't be made. */
function series(
  values: Value,
  timeline: Value,
  completion: Value | undefined,
  aggregation: Value | undefined,
  ctx: EvalContext,
): EtsSeries | CellError {
  const c = option(completion, 1, 0, 1)
  if (isError(c)) return c
  const g = option(aggregation, 1, 1, 7)
  if (isError(g)) return g
  const v = arrayArg(values, ctx)
  if (isError(v)) return v
  const t = arrayArg(timeline, ctx)
  if (isError(t)) return t
  if (v.data.length !== t.data.length) return NA
  const times: number[] = []
  const ys: (number | null)[] = []
  for (let i = 0; i < t.data.length; i++) {
    const tv: Scalar = t.data[i]
    const yv: Scalar = v.data[i]
    if (isError(tv)) return tv
    if (isError(yv)) return yv
    if (typeof tv !== "number") return VALUE
    if (yv !== null && typeof yv !== "number") return VALUE
    times.push(tv)
    ys.push(yv)
  }
  const s = etsSeries(times, ys, c, g)
  return typeof s === "string" ? problem(s) : s
}

/** The season length an argument asks for: 1 detects it, 0 means none. */
function seasonOf(v: Value | undefined, s: EtsSeries): number | CellError {
  const m = option(v, 1, 0, 8760)
  if (isError(m)) return m
  if (m === 1) return detectSeasonality(s.values)
  if (m > 1 && m * 2 > s.values.length) return NUM
  return m
}

/** Steps from the history's last point to the target; an error for a target before it. */
function horizon(target: Value, s: EtsSeries): number | CellError {
  const t = num(target)
  if (isError(t)) return t
  const h = etsPosition(s, t) - (s.values.length - 1)
  return h < -1e-9 ? NUM : Math.max(0, h)
}

define("FORECAST.ETS", {
  min: 3,
  max: 6,
  params: ["scalar", "any", "any", "scalar", "scalar", "scalar"],
  call: (a, ctx) => {
    const s = series(a[1], a[2], a[4], a[5], ctx)
    if (isError(s)) return s
    const m = seasonOf(a[3], s)
    if (isError(m)) return m
    const h = horizon(a[0], s)
    if (isError(h)) return h
    return etsForecast(fitEts(s.values, m), h)
  },
})

define("FORECAST.ETS.CONFINT", {
  min: 3,
  max: 7,
  params: ["scalar", "any", "any", "scalar", "scalar", "scalar", "scalar"],
  call: (a, ctx) => {
    let level = 0.95
    if (a[3] !== undefined && a[3] !== null) {
      const c = num(a[3])
      if (isError(c)) return c
      if (c <= 0 || c >= 1) return NUM
      level = c
    }
    const s = series(a[1], a[2], a[5], a[6], ctx)
    if (isError(s)) return s
    const m = seasonOf(a[4], s)
    if (isError(m)) return m
    const h = horizon(a[0], s)
    if (isError(h)) return h
    return etsConfidence(fitEts(s.values, m), h, level)
  },
})

define("FORECAST.ETS.SEASONALITY", {
  min: 2,
  max: 4,
  params: ["any", "any", "scalar", "scalar"],
  call: (a, ctx) => {
    const s = series(a[0], a[1], a[2], a[3], ctx)
    if (isError(s)) return s
    return detectSeasonality(s.values)
  },
})

define("FORECAST.ETS.STAT", {
  min: 3,
  max: 6,
  params: ["any", "any", "scalar", "scalar", "scalar", "scalar"],
  call: (a, ctx) => {
    const type = option(a[2], 0, 1, 8)
    if (isError(type)) return type
    const s = series(a[0], a[1], a[4], a[5], ctx)
    if (isError(s)) return s
    const m = seasonOf(a[3], s)
    if (isError(m)) return m
    return etsStatistic(fitEts(s.values, m), s, type) ?? NUM
  },
})
