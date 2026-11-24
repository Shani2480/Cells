import { describe, expect, it } from "vitest"
import {
  customStyle,
  customStyleNameProblem,
  duplicateStyleName,
  elementSummary,
  isEmptyElement,
  newCustomStyleName,
  STYLE_NAME_EMPTY,
  STYLE_NAME_TAKEN,
  type CustomTableStyle,
} from "./custom-styles"
import { applyElementFont, tableStyleDef, type TableCellLook } from "./styles"

const style = (name: string, kind: Partial<CustomTableStyle> = {}): CustomTableStyle => ({
  name,
  table: true,
  pivot: true,
  elements: { headerRow: { fill: "#112233" } },
  ...kind,
})

describe("custom table styles", () => {
  it("finds a style by name, ignoring case, among the right kind", () => {
    const list = [style("Mine", { table: false })]
    expect(customStyle(list, "mine", "pivot")?.name).toBe("Mine")
    expect(customStyle(list, "Mine", "table")).toBeUndefined()
    expect(tableStyleDef("Mine", undefined, list)).toBeNull()
    expect(tableStyleDef("Mine", undefined, [style("Mine")])?.headerRow?.fill).toBe("#112233")
  })

  it("names: the first free New name, Duplicate's numbered copy, and Excel's problems", () => {
    const list = [style("PivotTable Style 1"), style("Mine 2")]
    expect(newCustomStyleName(list)).toBe("PivotTable Style 2")
    expect(duplicateStyleName("Mine", list)).toBe("Mine 3")
    const builtIn = (n: string) => n === "PivotStyleLight16"
    expect(customStyleNameProblem(" ", list, builtIn)).toBe(STYLE_NAME_EMPTY)
    expect(customStyleNameProblem("pivottable style 1", list, builtIn)).toBe(STYLE_NAME_TAKEN)
    expect(customStyleNameProblem("PivotStyleLight16", list, builtIn)).toBe(STYLE_NAME_TAKEN)
    expect(customStyleNameProblem("PivotTable Style 1", list, builtIn, "PivotTable Style 1")).toBeNull()
  })

  it("an element's font turns flags on and, set to false, off again over the elements below", () => {
    const look: TableCellLook = { bold: true, italic: true }
    applyElementFont({ bold: false, strike: true, underline: "single", color: "#C00000" }, look)
    expect(look).toEqual({ italic: true, strike: true, underline: "single", fontColor: "#C00000" })
    expect(elementSummary({ bold: true, fill: "#fff", stripe: 2 })).toBe("Bold; Shaded; Stripe Size 2")
    expect(isEmptyElement({ stripe: 1, border: {} })).toBe(true)
    expect(isEmptyElement({ italic: false })).toBe(false)
  })
})
