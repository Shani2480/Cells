import { describe, expect, it } from "vitest"
import { MAX_COLS, MAX_ROWS, cellName, columnIndex, columnName, parseCell, parseRange, rangeName } from "../src/address"

describe("addresses", () => {
  it("names columns as Excel does", () => {
    expect([0, 25, 26, 51, 52, 701, 702, 16383].map(columnName)).toEqual([
      "A",
      "Z",
      "AA",
      "AZ",
      "BA",
      "ZZ",
      "AAA",
      "XFD",
    ])
    for (let c = 0; c < MAX_COLS; c += 97) expect(columnIndex(columnName(c))).toBe(c)
    expect(columnIndex("XFE")).toBe(-1)
    expect(columnIndex("A1")).toBe(-1)
  })

  it("parses cells and ranges from the Name Box", () => {
    expect(parseCell("b7")).toEqual({ row: 6, col: 1 })
    expect(parseCell("$B$7")).toEqual({ row: 6, col: 1 })
    expect(parseCell("A0")).toBeNull()
    expect(parseCell("A1048577")).toBeNull()
    expect(parseRange("C3:A1")).toEqual({ r0: 0, c0: 0, r1: 2, c1: 2 })
    expect(parseRange("B:D")).toEqual({ r0: 0, r1: MAX_ROWS - 1, c0: 1, c1: 3 })
    expect(parseRange("5:3")).toEqual({ r0: 2, r1: 4, c0: 0, c1: MAX_COLS - 1 })
    expect(parseRange("A1:B")).toBeNull()
    expect(parseRange("hello")).toBeNull()
  })

  it("prints ranges as Excel does", () => {
    expect(cellName({ row: 0, col: 0 })).toBe("A1")
    expect(rangeName({ r0: 0, c0: 0, r1: 2, c1: 1 })).toBe("A1:B3")
    expect(rangeName({ r0: 4, c0: 2, r1: 4, c1: 2 })).toBe("C5")
    expect(rangeName(parseRange("B:D")!)).toBe("B:D")
    expect(rangeName(parseRange("3:5")!)).toBe("3:5")
    expect(rangeName({ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 })).toBe("A1:XFD1048576")
  })
})
