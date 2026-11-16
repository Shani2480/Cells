import { describe, expect, it } from "vitest"
import { calculateFormula, calculateGrid, gridColumns, regionalFrom } from "../src"

describe("calculateGrid (Inkspire's Spreadsheet block)", () => {
  it("calculates formulas with Excel's functions, in any order", () => {
    const grid = calculateGrid([
      ["Item", "Cost", "", "=SUM(B2:B4)"],
      ["Rent", "1,200", "", "=AVERAGE(B2:B4)"],
      ["Food", "300.5", "", "=D1*2"],
      ["Bus", "10%", "", '=IF(D1>1000,"over","under")'],
    ])
    expect(grid[0]![3]).toEqual({ value: 1500.6, text: "1500.6" })
    expect(grid[2]![3]!.value).toBe(3001.2)
    expect(grid[3]![3]!.text).toBe("over")
    // Shown the way they were typed.
    expect(grid[1]![1]).toEqual({ value: 1200, text: "1,200" })
    expect(grid[3]![1]).toEqual({ value: 0.1, text: "10%" })
    expect(grid[0]![0]).toEqual({ value: "Item", text: "Item" })
    expect(grid[0]![2]).toEqual({ value: undefined, text: "" })
  })

  it("gives errors as Excel's codes, and a loop doesn't hang", () => {
    const grid = calculateGrid([["=1/0", "=NOSUCH(1)", "=A2", "=C1"]].concat([["=C1", "", "", ""]]))
    expect(grid[0]![0]).toMatchObject({ error: "#DIV/0!", text: "#DIV/0!" })
    expect(grid[0]![1]!.error).toBe("#NAME?")
    expect(grid[1]![0]!.text).toBeDefined()
  })

  it("reads and shows numbers and dates in the regional formats it is given", () => {
    const de = regionalFrom({ dateFormat: "dd.MM.yyyy", decimalSeparator: ",", groupSeparator: "." })
    const grid = calculateGrid([["1,5", "14.03.2026", "=A1*2", "=B1+1"]], { regional: de })
    expect(grid[0]![0]!.value).toBe(1.5)
    expect(grid[0]![2]!.text).toBe("3")
    expect(grid[0]![1]!.text).toBe("14.03.2026")
    expect(grid[0]![3]!.value).toBe(46096)
    // The default is untouched afterwards.
    expect(calculateGrid([["1,5"]])[0]![0]!.value).toBe("1,5")
  })

  it("names the columns", () => {
    expect(gridColumns(4)).toEqual(["A", "B", "C", "D"])
  })
})

describe("calculateFormula (Inkspire's formula fields)", () => {
  it("calculates a formula on its own", () => {
    expect(calculateFormula("=2+3*4")).toEqual({ ok: true, value: 14, text: "14" })
    expect(calculateFormula(" = ROUND(PI(), 2)")).toMatchObject({ ok: true, value: 3.14 })
    expect(calculateFormula('=TEXT(DATE(2026,1,15),"mmmm d")')).toMatchObject({ ok: true, text: "January 15" })
  })

  it("reads a grid's cells when given one", () => {
    expect(
      calculateFormula("=SUM(A1:B2)", [
        ["1", "2"],
        ["3", "4"],
      ]),
    ).toMatchObject({ value: 10 })
  })

  it("says why there is no result", () => {
    expect(calculateFormula("")).toEqual({ ok: false, error: "Invalid formula" })
    expect(calculateFormula("=2+*")).toEqual({ ok: false, error: "Invalid formula" })
    expect(calculateFormula("=1/0")).toEqual({ ok: false, error: "#DIV/0!" })
  })
})
