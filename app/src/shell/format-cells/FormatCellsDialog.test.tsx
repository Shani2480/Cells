import { parseRange, selectRange, writeInputOp } from "@cells/core"
import { act, fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { FormatCellsDialog } from "./FormatCellsDialog"
import { accountingCode, buildCode, currencyCode, numberCode, readCode } from "./number-formats"

function setup(tab?: "number" | "alignment" | "font" | "border" | "fill" | "protection") {
  const c = new SheetController()
  c.wb.transact([writeInputOp(0, c.sheet, c.wb.styles, 0, 0, "-1234.1")], { label: "t" })
  let closed = false
  render(<FormatCellsDialog controller={c} tab={tab} onClose={() => (closed = true)} />)
  const style = (ref: string) => {
    const r = parseRange(ref)!
    return c.wb.styles.get(c.sheet.styleAt(r.r0, r.c0))
  }
  return { c, style, closed: () => closed }
}

const ok = () => fireEvent.click(screen.getByRole("button", { name: "OK" }))

describe("number format codes", () => {
  it("builds Excel's codes for each category's options", () => {
    expect(numberCode({ decimals: 2, separator: true, negative: 3 })).toBe("#,##0.00_);[Red](#,##0.00)")
    expect(currencyCode({ decimals: 0, negative: 1, symbol: "$" })).toBe("$#,##0;[Red]$#,##0")
    expect(accountingCode({ decimals: 2, symbol: "$" })).toBe('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')
    expect(buildCode("Percentage", { decimals: 1, separator: false, negative: 0, symbol: "" }, undefined)).toBe("0.0%")
  })

  it("reads a code back into its category and options", () => {
    expect(readCode("#,##0.00_);[Red](#,##0.00)")).toMatchObject({
      category: "Number",
      options: { decimals: 2, separator: true, negative: 3 },
    })
    expect(readCode('_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)').category).toBe("Accounting")
    expect(readCode("0.000E+00")).toMatchObject({ category: "Scientific", options: { decimals: 3 } })
    expect(readCode("[h]:mm:ss;@").category).toBe("Time")
    expect(readCode('0.0" kg"').category).toBe("Custom")
  })
})

describe("Format Cells dialog", () => {
  it("Number: picking Number with a separator and red negatives shows a sample and applies the code", () => {
    const { style } = setup()
    fireEvent.click(screen.getByRole("option", { name: "Number" }))
    fireEvent.click(screen.getByLabelText("Use 1000 Separator (,)"))
    expect(screen.getByRole("group", { name: "Sample" }).textContent).toContain("-1,234.10")
    ok()
    expect(style("A1").numberFormat).toBe("#,##0.00")
  })

  it("applies only what changed, as one undo step", () => {
    const { c, style } = setup("font")
    act(() => c.setSelection(selectRange(parseRange("A1")!)))
    fireEvent.click(screen.getByRole("option", { name: "Bold Italic" }))
    ok()
    expect(style("A1")).toEqual({ bold: true, italic: true })
    c.undo()
    expect(style("A1")).toEqual({})
  })

  it("Cancel and Esc change nothing", () => {
    const { style, closed } = setup("font")
    fireEvent.click(screen.getByRole("option", { name: "Bold" }))
    fireEvent.click(screen.getByRole("button", { name: "Cancel" }))
    expect(closed()).toBe(true)
    expect(style("A1")).toEqual({})
  })

  it("Alignment: horizontal, wrap and merge", () => {
    const { c, style } = setup("alignment")
    act(() => c.setSelection(selectRange(parseRange("B2:C3")!)))
    fireEvent.change(screen.getByLabelText("Horizontal:"), { target: { value: "right" } })
    fireEvent.click(screen.getByLabelText("Wrap text"))
    fireEvent.click(screen.getByLabelText("Merge cells"))
    ok()
    expect(style("C3")).toMatchObject({ hAlign: "right", wrap: true })
    expect(c.sheet.merges.all).toEqual([parseRange("B2:C3")])
  })

  it("Border: Outline then Inside put lines round and through the selection", () => {
    const { c, style } = setup("border")
    act(() => c.setSelection(selectRange(parseRange("B2:C3")!)))
    fireEvent.click(screen.getByRole("button", { name: "Outline" }))
    fireEvent.click(screen.getByRole("button", { name: "Inside" }))
    ok()
    expect(style("B2")).toMatchObject({
      borderTop: { style: "thin" },
      borderLeft: { style: "thin" },
      borderBottom: { style: "thin" },
    })
    expect(style("C3")).toMatchObject({ borderBottom: { style: "thin" }, borderRight: { style: "thin" } })
  })

  it("Fill and Protection", () => {
    const { style } = setup("fill")
    fireEvent.click(screen.getAllByRole("menuitemradio", { name: "Yellow" })[0])
    fireEvent.click(screen.getByRole("tab", { name: "Protection" }))
    fireEvent.click(screen.getByLabelText("Locked"))
    ok()
    expect(style("A1")).toEqual({ fill: "#FFFF00", locked: false })
  })
})
