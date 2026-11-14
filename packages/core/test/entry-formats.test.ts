import { describe, expect, it } from "vitest"
import { parseDateTime } from "../src/entry-formats"
import { parseInput } from "../src/values"

// Every row was typed into Excel 16 (en-US) in 2026: [typed, value, number format] (undefined = General / text).
const EXCEL: [string, number | string, string | undefined][] = [
  ["$1,000", 1000, "$#,##0_);[Red]($#,##0)"],
  ["$1,234.50", 1234.5, "$#,##0.00_);[Red]($#,##0.00)"],
  ["-$5", -5, "$#,##0_);[Red]($#,##0)"],
  ["($5)", -5, "$#,##0_);[Red]($#,##0)"],
  ["$-5", -5, "$#,##0_);[Red]($#,##0)"],
  ["$ 5", 5, "$#,##0_);[Red]($#,##0)"],
  ["(5)", -5, undefined],
  ["1E5", 100000, "0.00E+00"],
  ["1.5e-3", 0.0015, "0.00E+00"],
  ["1 1/2", 1.5, "# ?/?"],
  ["0 1/2", 0.5, "# ?/?"],
  ["1/2", 46024, "d-mmm"],
  ["3/4", 46085, "d-mmm"],
  ["1/15/2026", 46037, "m/d/yyyy"],
  ["1/15/26", 46037, "m/d/yyyy"],
  ["15-Jan-2026", 46037, "d-mmm-yy"],
  ["15-Jan", 46037, "d-mmm"],
  ["Jan 15", 46037, "d-mmm"],
  ["Jan-2026", 46023, "mmm-yy"],
  ["January 15, 2026", 46037, "d-mmm-yy"],
  ["2026-01-15", 46037, "m/d/yyyy"],
  ["13:45", 0.572916666666667, "h:mm"],
  ["1:45 PM", 0.572916666666667, "h:mm AM/PM"],
  ["1:45:10 PM", 0.573032407407407, "h:mm:ss AM/PM"],
  ["13:45:10", 0.573032407407407, "h:mm:ss"],
  ["1/15/2026 13:45", 46037.5729166667, "m/d/yyyy h:mm"],
  ["1/15/2026 1:45 PM", 46037.5729166667, "m/d/yyyy h:mm"],
  ["25:30", 1.0625, "[h]:mm:ss"],
  ["10%", 0.1, "0%"],
  ["$10%", "$10%", undefined],
  ["1,000.5", 1000.5, "#,##0.00"],
  ["12/31", 46387, "d-mmm"],
  ["2/30/2026", "2/30/2026", undefined],
  ["5$", "5$", undefined],
  ["1e", "1e", undefined],
  ["12:00 AM", 0, "h:mm AM/PM"],
  ["0:05", 0.00347222222222222, "h:mm"],
  ["3:5", 0.128472222222222, "h:mm"],
  ["1/2/3", 37623, "m/d/yyyy"],
  ["Jan", "Jan", undefined],
  ["15 Jan 2026", 46037, "d-mmm-yy"],
  ["Jan 15 2026", "Jan 15 2026", undefined],
  ["mar-5", 46086, "d-mmm"],
  ["5/2026", 46143, "mmm-yy"],
  ["2026/5/15", 46157, "m/d/yyyy"],
]

describe("typed entries (as Excel reads them)", () => {
  it.each(EXCEL)("%s", (typed, value, format) => {
    // Dates typed without a year take the current one; the Excel run was in 2026.
    const parsed = typed.match(/^[A-Za-z]{3}[- ]\d{1,2}$|^\d{1,2}[-/][A-Za-z\d]{1,3}$|^mar-5$/i)
      ? { value: parseDateTime(typed, 2026)?.value, numberFormat: parseDateTime(typed, 2026)?.numberFormat }
      : parseInput(typed)!
    if (typeof value === "number") expect(parsed.value).toBeCloseTo(value, 9)
    else expect(parsed.value).toBe(value)
    expect(parsed.numberFormat).toBe(format)
  })
})
