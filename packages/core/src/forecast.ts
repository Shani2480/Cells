/**
 * Exponential triple smoothing (additive error, trend and seasonality: ETS AAA), as Excel's FORECAST.ETS functions
 * and Data › Forecast Sheet use it. A series is a timeline with a constant step (days, or whole months for dates on
 * the same day of the month) and its values; the model's smoothing factors are fitted to the history by least
 * squares of the one-step-ahead errors. Excel doesn't publish its fitting, so forecasts are close to Excel's, not
 * equal to them.
 */

/** Excel's aggregation codes for points on the same date: 1 AVERAGE (default) … 7 SUM. */
export const ETS_AGGREGATIONS = ["Average", "Count", "CountA", "Max", "Median", "Min", "Sum"] as const

/** A regular series: values at `start`, `start + step`, …; in months when `months`. */
export type EtsSeries = {
  values: number[]
  start: number
  step: number
  months: boolean
  /** Monthly dates all on the last day of their month. */
  monthEnds?: boolean
}

export type EtsProblem = "num" | "na" | "value"

/** Serial date → whole months since year 0 and the day of the month (0 for the last day). */
function monthOf(serial: number): { month: number; day: number; last: boolean } {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000)
  const day = d.getUTCDate()
  const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate() === day
  return { month: d.getUTCFullYear() * 12 + d.getUTCMonth(), day, last }
}

/** Month index (as monthOf) and day → serial date; days past the month's end fall on its last day. */
export function serialOfMonth(month: number, day: number, last: boolean): number {
  const y = Math.floor(month / 12)
  const m = month - y * 12
  const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate()
  const d = last ? lastDay : Math.min(day, lastDay)
  return Math.round((Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000)
}

function aggregate(code: number, xs: number[], count: number): number {
  switch (code) {
    case 2:
      return xs.length
    case 3:
      return count
    case 4:
      return Math.max(...xs)
    case 5: {
      const s = [...xs].sort((a, b) => a - b)
      const h = s.length >> 1
      return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2
    }
    case 6:
      return Math.min(...xs)
    case 7:
      return xs.reduce((a, b) => a + b, 0)
    default:
      return xs.reduce((a, b) => a + b, 0) / xs.length
  }
}

/**
 * The regular series from a timeline and its values (null: a missing value). Points on the same date are combined
 * with `aggregation`; missing points are filled in a straight line (`completion` 1) or with zeros (0). A problem
 * when the timeline has no constant step (#NUM!) or too few points.
 */
export function etsSeries(
  timeline: number[],
  values: (number | null)[],
  completion = 1,
  aggregation = 1,
): EtsSeries | EtsProblem {
  if (timeline.length !== values.length) return "na"
  const groups = new Map<number, { xs: number[]; count: number }>()
  timeline.forEach((t, i) => {
    const g = groups.get(t) ?? { xs: [], count: 0 }
    const v = values[i]
    if (v !== null) {
      g.xs.push(v)
      g.count++
    }
    groups.set(t, g)
  })
  const times = [...groups.keys()].sort((a, b) => a - b)
  if (times.length < 2) return "num"
  const points = new Map<number, number | null>()
  for (const t of times) {
    const g = groups.get(t)!
    points.set(t, g.xs.length || aggregation === 2 || aggregation === 3 ? aggregate(aggregation, g.xs, g.count) : null)
  }
  // Whole months: dates on the same day of the month (or all on the last day) with months between them.
  let months = false
  let monthEnds = false
  let index = (t: number) => t
  if (times.every((t) => Number.isInteger(t) && t > 0)) {
    const ms = times.map(monthOf)
    const sameDay = ms.every((m) => m.day === ms[0].day) || ms.every((m) => m.last)
    const gaps = new Set(times.slice(1).map((t, i) => t - times[i]))
    if (sameDay && gaps.size > 1 && ms.every((m, i) => i === 0 || m.month > ms[i - 1].month)) {
      months = true
      monthEnds = !ms.every((m) => m.day === ms[0].day)
      index = (t: number) => monthOf(t).month
    }
  }
  const at = times.map(index)
  let step = Infinity
  for (let i = 1; i < at.length; i++) step = Math.min(step, at[i] - at[i - 1])
  if (!(step > 0)) return "num"
  const n = Math.round((at[at.length - 1] - at[0]) / step) + 1
  if (n > 100000) return "num"
  const regular: (number | null)[] = new Array(n).fill(null)
  for (let i = 0; i < at.length; i++) {
    const k = (at[i] - at[0]) / step
    if (Math.abs(k - Math.round(k)) > 1e-7) return "num"
    regular[Math.round(k)] = points.get(times[i]) ?? null
  }
  // Fill the missing points.
  const filled: number[] = []
  for (let i = 0; i < n; i++) {
    const v = regular[i]
    if (v !== null) {
      filled.push(v)
      continue
    }
    if (completion === 0) {
      filled.push(0)
      continue
    }
    let a = i - 1
    while (a >= 0 && regular[a] === null) a--
    let b = i + 1
    while (b < n && regular[b] === null) b++
    const va = a >= 0 ? regular[a]! : null
    const vb = b < n ? regular[b]! : null
    filled.push(va === null ? (vb ?? 0) : vb === null ? va : va + ((vb - va) * (i - a)) / (b - a))
  }
  return { values: filled, start: months ? times[0] : at[0], step, months, ...(monthEnds ? { monthEnds } : {}) }
}

/** Where a timeline value falls on the series, in steps from its start. */
export function etsPosition(series: EtsSeries, t: number): number {
  if (!series.months) return (t - series.start) / series.step
  const a = monthOf(series.start).month
  const b = monthOf(t)
  return (b.month - a) / series.step
}

/** The timeline value `k` steps from the series' start. */
export function etsTime(series: EtsSeries, k: number): number {
  if (!series.months) return series.start + k * series.step
  const m = monthOf(series.start)
  return serialOfMonth(m.month + k * series.step, m.day, series.monthEnds === true)
}

const SEASON_THRESHOLD = 0.25

/** The season length seen in the values (autocorrelation of what's left after the trend); 1 when there's none. */
export function detectSeasonality(y: number[]): number {
  const n = y.length
  if (n < 6) return 1
  // Remove a straight-line trend.
  const mx = (n - 1) / 2
  const my = y.reduce((a, b) => a + b, 0) / n
  let sxy = 0
  let sxx = 0
  for (let i = 0; i < n; i++) {
    sxy += (i - mx) * (y[i] - my)
    sxx += (i - mx) ** 2
  }
  const slope = sxy / sxx
  const r = y.map((v, i) => v - my - slope * (i - mx))
  const c0 = r.reduce((a, b) => a + b * b, 0)
  if (c0 < 1e-12) return 1
  const max = Math.floor(n / 2)
  const acf = [1]
  for (let k = 1; k <= max + 1; k++) {
    let s = 0
    for (let i = k; i < n; i++) s += r[i] * r[i - k]
    acf.push(s / c0)
  }
  // The first peak strong enough, as Excel picks it (a pattern of 4 that also echoes at 2 is reported as 2).
  for (let k = 2; k <= max; k++) if (acf[k] > acf[k - 1] && acf[k] >= acf[k + 1] && acf[k] > SEASON_THRESHOLD) return k
  return 1
}

export type EtsModel = {
  alpha: number
  beta: number
  gamma: number
  /** Season length; 0 without seasonality. */
  m: number
  level: number
  trend: number
  /** The last season's factors, oldest first. */
  season: number[]
  /** One-step-ahead predictions of the history (index i predicts y[i]). */
  fitted: number[]
  y: number[]
  sse: number
}

function run(y: number[], m: number, alpha: number, beta: number, gamma: number, keep = false) {
  const n = y.length
  // The starting state from the whole history: a line through the means of its whole seasons (so a season's
  // pattern doesn't tilt it), and each position's average offset from that line. Without seasons, a least-squares
  // line through every point. Close to the state Excel starts from.
  let level: number
  let trend: number
  let season: number[] = []
  const line = (xs: number[], ys: number[]) => {
    const k = xs.length
    const mx = xs.reduce((a, b) => a + b, 0) / k
    const my = ys.reduce((a, b) => a + b, 0) / k
    let sxy = 0
    let sxx = 0
    for (let i = 0; i < k; i++) {
      sxy += (xs[i] - mx) * (ys[i] - my)
      sxx += (xs[i] - mx) ** 2
    }
    const slope = sxx ? sxy / sxx : 0
    return { slope, intercept: my - slope * mx }
  }
  if (m > 1) {
    const cycles = Math.floor(n / m)
    const xs: number[] = []
    const ys: number[] = []
    for (let c = 0; c < cycles; c++) {
      xs.push(c * m + (m - 1) / 2)
      ys.push(y.slice(c * m, c * m + m).reduce((a, b) => a + b, 0) / m)
    }
    const { slope, intercept } = line(xs, ys)
    const sums = new Array<number>(m).fill(0)
    for (let i = 0; i < cycles * m; i++) sums[i % m] += y[i] - intercept - slope * i
    season = sums.map((sum) => sum / cycles)
    const avg = season.reduce((a, b) => a + b, 0) / m
    season = season.map((v) => v - avg)
    trend = slope
    level = intercept - slope
  } else {
    const { slope, intercept } = line(
      y.map((_, i) => i),
      y,
    )
    trend = slope
    level = intercept - slope
  }
  let sse = 0
  const fitted: number[] = keep ? [] : (undefined as unknown as number[])
  for (let t = 0; t < n; t++) {
    const s = m > 1 ? season[t % m] : 0
    const predicted = level + trend + s
    const e = y[t] - predicted
    sse += e * e
    if (keep) fitted.push(predicted)
    const previous = level
    level = alpha * (y[t] - s) + (1 - alpha) * (level + trend)
    trend = beta * (level - previous) + (1 - beta) * trend
    if (m > 1) season[t % m] = gamma * (y[t] - level) + (1 - gamma) * s
  }
  // The season rotated so it starts at the point after the history.
  const rotated = m > 1 ? season.slice(n % m).concat(season.slice(0, n % m)) : []
  return { sse, level, trend, season: rotated, fitted }
}

const fits = new Map<string, EtsModel>()

/** The model fitted to `y` with season length `m` (0 or 1: none). */
export function fitEts(y: number[], m: number): EtsModel {
  const season = m > 1 ? m : 0
  const key = `${season}|${y.join(",")}`
  const cached = fits.get(key)
  if (cached) return cached
  // Excel's starting point, which it keeps when nothing fits better (alpha 0.751, beta and gamma 0.001).
  const startAt = { a: 0.751, b: 0.001, g: season ? 0.001 : 0 }
  let best = { ...startAt, sse: run(y, season, startAt.a, startAt.b, startAt.g).sse }
  const better = (sse: number) => sse < best.sse - 1e-9 * Math.max(1, best.sse)
  // A local search from there, one factor at a time with shrinking steps (Excel's search is local too: a history
  // with no better fit nearby keeps the starting factors).
  const clamp = (v: number) => Math.min(0.999, Math.max(0.001, v))
  for (let step = 0.25; step > 1e-4; step /= 2) {
    let moved = true
    while (moved) {
      moved = false
      for (const key of season ? (["a", "b", "g"] as const) : (["a", "b"] as const))
        for (const dir of [-1, 1]) {
          const next = { ...best, [key]: clamp(best[key] + dir * step) }
          const { sse } = run(y, season, next.a, next.b, next.g)
          if (better(sse)) {
            best = { ...next, sse }
            moved = true
          }
        }
    }
  }
  const final = run(y, season, best.a, best.b, best.g, true)
  const model: EtsModel = {
    alpha: best.a,
    beta: best.b,
    gamma: season ? best.g : 0,
    m: season,
    level: final.level,
    trend: final.trend,
    season: final.season,
    fitted: final.fitted,
    y,
    sse: final.sse,
  }
  if (fits.size > 16) fits.delete(fits.keys().next().value!)
  fits.set(key, model)
  return model
}

/** The forecast `h` steps after the history's last point (h may be a fraction; 0 is the smoothed last point). */
export function etsForecast(model: EtsModel, h: number): number {
  const steps = Math.max(0, h)
  const s = model.m ? model.season[(((Math.ceil(steps) - 1) % model.m) + model.m) % model.m] : 0
  return model.level + steps * model.trend + s
}

/** The half-width of the confidence interval `h` steps ahead, at `confidence` (0–1). */
export function etsConfidence(model: EtsModel, h: number, confidence: number): number {
  const n = model.y.length
  const sigma2 = model.sse / n
  let sum = 1
  const steps = Math.ceil(Math.max(0, h))
  for (let j = 1; j < steps; j++) {
    const c = model.alpha * (1 + j * model.beta) + (model.m && j % model.m === 0 ? model.gamma : 0)
    sum += c * c
  }
  return normalInverse((1 + confidence) / 2) * Math.sqrt(sigma2 * sum)
}

/** FORECAST.ETS.STAT's statistics: 1 alpha … 8 step size. */
export function etsStatistic(model: EtsModel, series: EtsSeries, type: number): number | null {
  const y = model.y
  const n = y.length
  const errors = y.map((v, i) => v - model.fitted[i])
  switch (type) {
    case 1:
      return model.alpha
    case 2:
      return model.beta
    case 3:
      return model.gamma
    case 4: {
      let naive = 0
      for (let i = 1; i < n; i++) naive += Math.abs(y[i] - y[i - 1])
      naive /= n - 1
      const mae = errors.reduce((a, e) => a + Math.abs(e), 0) / n
      return naive ? mae / naive : null
    }
    case 5:
      return (
        errors.reduce((a, e, i) => {
          const d = (Math.abs(y[i]) + Math.abs(model.fitted[i])) / 2
          return a + (d ? Math.abs(e) / d : 0)
        }, 0) / n
      )
    case 6:
      return errors.reduce((a, e) => a + Math.abs(e), 0) / n
    case 7:
      return Math.sqrt(errors.reduce((a, e) => a + e * e, 0) / n)
    case 8:
      return series.step
    default:
      return null
  }
}

/** The standard normal distribution's inverse (Acklam's approximation, to about 1e-9). */
export function normalInverse(p: number): number {
  if (p <= 0) return -Infinity
  if (p >= 1) return Infinity
  const a = [
    -39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239,
  ]
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572]
  const c = [
    -0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968,
    2.938163982698783,
  ]
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416]
  const low = 0.02425
  let x: number
  if (p < low) {
    const q = Math.sqrt(-2 * Math.log(p))
    x =
      (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  } else if (p <= 1 - low) {
    const q = p - 0.5
    const r = q * q
    x =
      ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) /
      (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p))
    x =
      -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  }
  return x
}
