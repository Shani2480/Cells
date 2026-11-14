import { describe, expect, it } from "vitest"
import { editText, errorValue, parseInput } from "../src/values"

describe("typed input", () => {
  it.each([
    ["42", 42],
    ["-3.5", -3.5],
    [".5", 0.5],
    ["1e3", 1000],
    ["  7 ", 7],
    ["007", 7],
    ["+4", 4],
    ["1,234", 1234],
    ["1,234.5", 1234.5],
  ])("reads %j as the number %j", (text, value) => {
    expect(parseInput(text)!.value).toBe(value)
    expect(parseInput(text)!.input).toBeUndefined()
  })

  it("keeps text that is not a number", () => {
    for (const text of ["1,23", "12a", "1.2.3", "abc", "- 4", "1 000"]) {
      expect(parseInput(text)!.value).toBe(text)
    }
  })

  it("sets the entry formats Excel sets", () => {
    expect(parseInput("10%")).toEqual({ value: 0.1, numberFormat: "0%" })
    expect(parseInput("12.50%")).toEqual({ value: 0.125, numberFormat: "0.00%" })
    expect(parseInput("1,000")).toEqual({ value: 1000, numberFormat: "#,##0" })
    expect(parseInput("1,000.25")).toEqual({ value: 1000.25, numberFormat: "#,##0.00" })
  })

  it("reads booleans, errors, apostrophes and formulas", () => {
    expect(parseInput("true")!.value).toBe(true)
    expect(parseInput("FALSE")!.value).toBe(false)
    expect(parseInput("#n/a")!.value).toBe(errorValue("#N/A"))
    expect(parseInput("'123")).toEqual({ value: "123", input: "'123" })
    expect(parseInput("=A1+1")).toEqual({ value: "", input: "=A1+1", formula: true })
    expect(parseInput("=")!.value).toBe("=")
    expect(parseInput("")).toBeNull()
  })

  it("keeps 15 significant digits, cutting off the rest as Excel does", () => {
    expect(parseInput("1234567890123456789")!.value).toBe(1234567890123450000)
    expect(parseInput("0.1234567890123456")!.value).toBe(0.123456789012345)
  })

  it("gives the formula bar text of a value", () => {
    expect(editText(0.1 + 0.2)).toBe("0.3")
    expect(editText(1e21)).toBe("1E+21")
    expect(editText(1.5e-10)).toBe("1.5E-10")
    expect(editText(123456789012345)).toBe("123456789012345")
    expect(editText(true)).toBe("TRUE")
    expect(editText(errorValue("#DIV/0!"))).toBe("#DIV/0!")
  })
})
