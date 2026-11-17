import { describe, expect, it } from "vitest"
import {
  calculateFormula,
  detectSeasonality,
  etsForecast,
  etsSeries,
  etsTime,
  fitEts,
  normalInverse,
  type EtsSeries,
} from "../src"

const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000

describe("ETS series", () => {
  it("fills missing points in a line or with zeros, and combines repeated dates", () => {
    const s = etsSeries([1, 2, 4, 4, 5], [10, 20, 40, 60, 60], 1, 1) as EtsSeries
    expect(s).toEqual({ values: [10, 20, 35, 50, 60], start: 1, step: 1, months: false })
    expect((etsSeries([1, 2, 4], [1, 2, 4], 0) as EtsSeries).values).toEqual([1, 2, 0, 4])
    expect((etsSeries([3, 1, 2, 2], [3, 1, 1, 3], 1, 7) as EtsSeries).values).toEqual([1, 4, 3])
  })

  it("needs a constant step and matching lengths", () => {
    expect(etsSeries([1, 2, 3.5], [1, 2, 3])).toBe("num")
    expect(etsSeries([1, 2], [1])).toBe("na")
  })

  it("steps in whole months for dates on the same day of the month or at month ends", () => {
    const s = etsSeries([serial(2026, 1, 15), serial(2026, 2, 15), serial(2026, 3, 15)], [1, 2, 3]) as EtsSeries
    expect(s.months).toBe(true)
    expect(etsTime(s, 3)).toBe(serial(2026, 4, 15))
    const ends = etsSeries([serial(2026, 1, 31), serial(2026, 2, 28), serial(2026, 3, 31)], [1, 2, 3]) as EtsSeries
    expect(etsTime(ends, 3)).toBe(serial(2026, 4, 30))
    // Weekly dates stay in days.
    expect(
      (etsSeries([serial(2026, 1, 1), serial(2026, 1, 8), serial(2026, 1, 15)], [1, 2, 3]) as EtsSeries).step,
    ).toBe(7)
  })
})

describe("ETS model", () => {
  it("finds a repeating pattern as Excel does (its first strong echo), and none in a straight line", () => {
    const series = (season: number[]) => Array.from({ length: 24 }, (_, i) => 100 + 2 * i + season[i % 4])
    expect(detectSeasonality(series([5, -3, -8, 6]))).toBe(4)
    // A pattern of 4 that also repeats every 2 (5 and 8 high, -3 and -10 low): Excel reports 2.
    expect(detectSeasonality(series([5, -3, 8, -10]))).toBe(2)
    expect(detectSeasonality(Array.from({ length: 24 }, (_, i) => 3 * i + 1))).toBe(1)
  })

  it("continues a straight line, and a trend with seasons", () => {
    const line = fitEts(
      Array.from({ length: 12 }, (_, i) => 10 + 5 * i),
      0,
    )
    expect(etsForecast(line, 1)).toBeCloseTo(70, 6)
    expect(etsForecast(line, 3)).toBeCloseTo(80, 6)
    const season = [5, -3, 8, -10]
    const y = Array.from({ length: 24 }, (_, i) => 100 + 2 * i + season[i % 4])
    const model = fitEts(y, 4)
    for (let h = 1; h <= 8; h++) expect(etsForecast(model, h)).toBeCloseTo(100 + 2 * (23 + h) + season[(23 + h) % 4], 1)
  })

  it("inverts the normal distribution", () => {
    expect(normalInverse(0.975)).toBeCloseTo(1.959964, 6)
    expect(normalInverse(0.5)).toBeCloseTo(0, 9)
    expect(normalInverse(0.01)).toBeCloseTo(-2.326348, 6)
  })
})

describe("FORECAST.ETS functions", () => {
  const cells = Array.from({ length: 12 }, (_, i) => [String(i + 1), String(10 + 5 * i + (i % 2 ? 1 : -1))])
  const value = (f: string) => calculateFormula(f, cells)

  it("forecast, confidence interval, seasonality and statistics", () => {
    const f = value("FORECAST.ETS(13,B1:B12,A1:A12,0)")
    expect(f.ok && Math.abs((f.value as number) - 70) < 2).toBe(true)
    const near = value("FORECAST.ETS.CONFINT(13,B1:B12,A1:A12,0.95,0)")
    const far = value("FORECAST.ETS.CONFINT(20,B1:B12,A1:A12,0.95,0)")
    expect(near.ok && far.ok && (far.value as number) > (near.value as number) && (near.value as number) > 0).toBe(true)
    expect(value("FORECAST.ETS.STAT(B1:B12,A1:A12,8)")).toMatchObject({ ok: true, value: 1 })
    expect(value("FORECAST.ETS.SEASONALITY(B1:B12,A1:A12)")).toMatchObject({ ok: true })
  })

  it("Excel's errors", () => {
    expect(value("FORECAST.ETS(5,B1:B12,A1:A12)")).toEqual({ ok: false, error: "#NUM!" })
    expect(value("FORECAST.ETS(13,B1:B12,A1:A11)")).toEqual({ ok: false, error: "#N/A" })
    expect(value("FORECAST.ETS(13,B1:B12,A1:A12,1,2)")).toEqual({ ok: false, error: "#NUM!" })
    expect(value("FORECAST.ETS.CONFINT(13,B1:B12,A1:A12,1)")).toEqual({ ok: false, error: "#NUM!" })
    expect(value("FORECAST.ETS.STAT(B1:B12,A1:A12,9)")).toEqual({ ok: false, error: "#NUM!" })
  })
})
