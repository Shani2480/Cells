import { afterEach, describe, expect, it } from "vitest"
import {
  EN_US,
  displayText,
  editTextFor,
  formatGeneral,
  formatWithCode,
  parseInput,
  regionalFrom,
  setRegional,
  withRegional,
  Workbook,
} from "../src"

/** A formula's result, calculated in a workbook. */
function evaluateFormula(formula: string): unknown {
  const wb = new Workbook(["Sheet1"])
  wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: parseInput(formula, EN_US)! }] }], {
    label: "t",
  })
  return wb.sheets[0].cells.value(0, 0)
}

const DE = regionalFrom({ dateFormat: "dd.MM.yyyy", decimalSeparator: ",", groupSeparator: "." })
const GB = regionalFrom({ dateFormat: "dd/MM/yyyy", decimalSeparator: ".", groupSeparator: "," })
const ISO = regionalFrom({ dateFormat: "yyyy-MM-dd", decimalSeparator: ",", groupSeparator: " " })

afterEach(() => setRegional(EN_US))

const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000

describe("the regional setting from the shared settings", () => {
  it("reads the order, separator and short date from the pattern", () => {
    expect(DE).toEqual({ decimal: ",", group: ".", dateSeparator: ".", dateOrder: "dmy", shortDate: "dd/mm/yyyy" })
    expect(GB).toMatchObject({ dateOrder: "dmy", dateSeparator: "/", shortDate: "dd/mm/yyyy", decimal: "." })
    expect(ISO).toMatchObject({ dateOrder: "ymd", dateSeparator: "-", shortDate: "yyyy/mm/dd", group: " " })
    expect(regionalFrom({})).toEqual(EN_US)
  })
})

describe("showing numbers and dates", () => {
  it("uses the region's separators in number formats and General", () => {
    setRegional(DE)
    expect(formatWithCode(1234567.891, "#,##0.00").text).toBe("1.234.567,89")
    expect(formatWithCode(0.125, "0.0%").text).toBe("12,5%")
    expect(formatWithCode(12345, "0.00E+00").text).toBe("1,23E+04")
    expect(formatWithCode(-5, '"x."0.0').text).toBe("-x.5,0")
    expect(formatGeneral(1.5)).toBe("1,5")
    expect(formatGeneral(123456789012)).toBe("1,23457E+11")
  })

  it("shows the built-in short date in the region's order, and / as its separator", () => {
    const d = serial(2026, 1, 31)
    setRegional(DE)
    expect(displayText(d, "m/d/yyyy")).toBe("31.01.2026")
    expect(displayText(d + 0.5, "m/d/yyyy h:mm")).toBe("31.01.2026 12:00")
    expect(displayText(d, "d/m/yy")).toBe("31.1.26")
    expect(displayText(d, "d-mmm-yy")).toBe("31-Jan-26")
    setRegional(GB)
    expect(displayText(d, "m/d/yyyy")).toBe("31/01/2026")
    setRegional(ISO)
    expect(displayText(d, "m/d/yyyy")).toBe("2026-01-31")
  })

  it("puts the formula bar's numbers and dates in the region's form", () => {
    setRegional(DE)
    expect(editTextFor(1.5)).toBe("1,5")
    expect(editTextFor(0.125, "0.0%")).toBe("12,5%")
    expect(editTextFor(serial(2026, 1, 31), "d-mmm-yy")).toBe("31.01.2026")
  })
})

describe("reading typed entries", () => {
  it("reads decimal commas and dot thousands in a German setting", () => {
    setRegional(DE)
    expect(parseInput("1,5")).toEqual({ value: 1.5 })
    expect(parseInput("1.234")).toEqual({ value: 1234, numberFormat: "#,##0" })
    expect(parseInput("1.234,5")).toEqual({ value: 1234.5, numberFormat: "#,##0.00" })
    expect(parseInput("12,5%")).toEqual({ value: 0.125, numberFormat: "0.0%" })
    expect(parseInput("-1,5E3")).toMatchObject({ value: -1500 })
    expect(parseInput("$1.000,50")).toEqual({ value: 1000.5, numberFormat: "$#,##0.00_);[Red]($#,##0.00)" })
  })

  it("reads dates day first, with the region's separator", () => {
    setRegional(DE)
    expect(parseInput("31.1.2026")).toEqual({ value: serial(2026, 1, 31), numberFormat: "m/d/yyyy" })
    // "1.5" is the first of May in Germany, as Excel there reads it.
    expect(parseInput("1.5")).toEqual({ value: serial(new Date().getFullYear(), 5, 1), numberFormat: "d-mmm" })
    expect(parseInput("31/1/2026 13:45")).toMatchObject({ value: serial(2026, 1, 31) + 13.75 / 24 })
    setRegional(GB)
    expect(parseInput("31/1/2026")).toEqual({ value: serial(2026, 1, 31), numberFormat: "m/d/yyyy" })
    expect(parseInput("15/1")).toMatchObject({ value: serial(new Date().getFullYear(), 1, 15) })
    expect(parseInput("1/15/2026")).toEqual({ value: "1/15/2026" })
    expect(parseInput("1.5")).toEqual({ value: 1.5 })
  })

  it("reads year-first dates and space thousands", () => {
    setRegional(ISO)
    expect(parseInput("2026-1-31")).toMatchObject({ value: serial(2026, 1, 31) })
    expect(parseInput("1 234,5")).toEqual({ value: 1234.5, numberFormat: "#,##0.00" })
  })

  it("keeps en-US rules where asked, and puts the setting back", () => {
    setRegional(DE)
    expect(withRegional(EN_US, () => parseInput("1.5"))).toEqual({ value: 1.5 })
    expect(parseInput("1.5", EN_US)).toEqual({ value: 1.5 })
    expect(parseInput("1,5")).toEqual({ value: 1.5 })
  })

  it("leaves formulas in Excel's en-US form: TEXT formats as en-US Excel", () => {
    setRegional(DE)
    expect(evaluateFormula('=TEXT(1234.5,"#,##0.00")')).toBe("1,234.50")
    expect(evaluateFormula('=VALUE("1.5")')).toBe(1.5)
  })
})
