import { describe, expect, it } from "vitest"
import { formatGeneral, formatValue } from "../src/format"

describe("General format in a default-width column", () => {
  // Expectations from Excel's default 8.43-character column (to be re-checked against Excel's TEXT/display in Phase 3).
  it.each([
    [0, "0"],
    [42, "42"],
    [-42, "-42"],
    [1 / 3, "0.333333333"],
    [2 / 3, "0.666666667"],
    [-1 / 3, "-0.33333333"],
    [123.456789012, "123.456789"],
    [1234567890, "1234567890"],
    [12345678901, "12345678901"],
    [123456789012, "1.23457E+11"],
    [0.1 + 0.2, "0.3"],
    [0.00001, "0.00001"],
    [1e-10, "1E-10"],
    [1.5e20, "1.5E+20"],
  ])("%d → %s", (value, text) => {
    expect(formatGeneral(value)).toBe(text)
  })

  it("shortens decimals to fit narrow columns, then gives up", () => {
    expect(formatGeneral(1 / 3, 5)).toBe("0.333")
    expect(formatGeneral(123456, 5)).toBe("1E+05")
    expect(formatGeneral(123456, 3)).toBeNull()
  })

  it("draws entry formats", () => {
    expect(formatValue(0.1, "0%")).toBe("10%")
    expect(formatValue(0.125, "0.00%")).toBe("12.50%")
    expect(formatValue(1234567, "#,##0")).toBe("1,234,567")
    expect(formatValue(-1234.5, "#,##0.00")).toBe("-1,234.50")
    expect(formatValue(true)).toBe("TRUE")
  })
})
