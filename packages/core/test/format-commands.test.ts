import { describe, expect, it } from "vitest"
import { parseRange } from "../src/address"
import { borderOps, changeDecimals, stepFontSize } from "../src/format-commands"
import { Workbook } from "../src/workbook"

describe("Increase/Decrease Decimal", () => {
  // Expectations recorded from Excel's own Increase/Decrease Decimal (CommandBars.ExecuteMso in the VM).
  it("on General starts from what the cell shows", () => {
    expect(changeDecimals(undefined, 1, "3.14159")).toBe("0.000000")
    expect(changeDecimals("General", -1, "3.14159")).toBe("0.0000")
    expect(changeDecimals(undefined, 1, "12")).toBe("0.0")
    expect(changeDecimals(undefined, -1, "12")).toBeNull()
    expect(changeDecimals(undefined, 1, "1234567.9")).toBe("0.00")
    expect(changeDecimals(undefined, -1, "0.000123")).toBe("0.00000")
  })

  it("adds or removes a 0 in every number section", () => {
    expect(changeDecimals("0.00", 1)).toBe("0.000")
    expect(changeDecimals("0.0", -1)).toBe("0")
    expect(changeDecimals("0", 1)).toBe("0.0")
    expect(changeDecimals("#,##0.00;[Red]-#,##0.00", -1)).toBe("#,##0.0;[Red]-#,##0.0")
    expect(changeDecimals("0%", 1)).toBe("0.0%")
    expect(changeDecimals('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', -1)).toBe(
      '_($* #,##0.0_);_($* (#,##0.0);_($* "-"??_);_(@_)',
    )
    expect(changeDecimals("0.00E+00", 1)).toBe("0.000E+00")
    expect(changeDecimals("0.00E+00", -1)).toBe("0.0E+00")
    expect(changeDecimals('"x"0.0"y"', -1)).toBe('"x"0"y"')
    expect(changeDecimals("$#,##0", 1)).toBe("$#,##0.0")
    expect(changeDecimals("0", -1)).toBeNull()
    expect(changeDecimals("m/d/yyyy", 1)).toBeNull()
    expect(changeDecimals("# ?/?", 1)).toBeNull()
  })
})

describe("Grow and Shrink Font", () => {
  it("steps through Excel's size list", () => {
    expect(stepFontSize(11, 1)).toBe(12)
    expect(stepFontSize(12, 1)).toBe(14)
    expect(stepFontSize(13, -1)).toBe(12)
    expect(stepFontSize(72, 1)).toBe(80)
    expect(stepFontSize(8, -1)).toBe(7)
  })
})

describe("border presets", () => {
  const apply = (preset: Parameters<typeof borderOps>[2], ref = "B2:D4") => {
    const wb = new Workbook(["Sheet1"])
    wb.transact(borderOps(0, parseRange(ref)!, preset), { label: "Borders" })
    const at = (cell: string) => {
      const { r0, c0 } = parseRange(cell)!
      const s = wb.styles.get(wb.sheets[0].styleAt(r0, c0))
      return ["borderTop", "borderBottom", "borderLeft", "borderRight"].filter((k) => k in s).join(",")
    }
    return { wb, at }
  }

  it("Outside Borders frame the range", () => {
    const { at } = apply("outside")
    expect(at("B2")).toBe("borderTop,borderLeft")
    expect(at("C3")).toBe("")
    expect(at("D4")).toBe("borderBottom,borderRight")
  })

  it("All Borders give every cell four edges; No Border takes them and the neighbours' facing edges away", () => {
    const { wb, at } = apply("all")
    expect(at("C3")).toBe("borderTop,borderBottom,borderLeft,borderRight")
    wb.transact(borderOps(0, parseRange("A1:A1")!, "bottom"), { label: "B" })
    wb.transact(borderOps(0, parseRange("A2:A2")!, "none"), { label: "None" })
    expect(at("A1")).toBe("")
  })

  it("Thick Bottom and Top and Double Bottom", () => {
    const { wb } = apply("topDoubleBottom")
    const s = (cell: string) => {
      const { r0, c0 } = parseRange(cell)!
      return wb.styles.get(wb.sheets[0].styleAt(r0, c0))
    }
    expect(s("C2").borderTop).toEqual({ style: "thin" })
    expect(s("C4").borderBottom).toEqual({ style: "double" })
  })
})
