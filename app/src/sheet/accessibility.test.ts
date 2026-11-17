import { parseRange, selectRange } from "@cells/core"
import { describe, expect, it } from "vitest"
import { checkAccessibility, contrastRatio, issueActions } from "./accessibility"
import { SheetController } from "./controller"

function setup() {
  const c = new SheetController()
  c.wb.transact(
    [
      {
        kind: "setCells",
        sheet: 0,
        cells: [
          { row: 0, col: 0, data: { value: "Month" } },
          { row: 0, col: 1, data: { value: "Sales" } },
          { row: 1, col: 0, data: { value: "Jan" } },
          { row: 1, col: 1, data: { value: 40 } },
          { row: 2, col: 0, data: { value: "Feb" } },
          { row: 2, col: 1, data: { value: 55 } },
        ],
      },
    ],
    { label: "x", before: c.selection },
  )
  return c
}

const kinds = (c: SheetController) => checkAccessibility(c.wb).map((i) => i.kind)

describe("Check Accessibility", () => {
  it("measures contrast as WCAG does", () => {
    expect(contrastRatio("#000000", "#ffffff")).toBeCloseTo(21)
    expect(contrastRatio("#777777", "#ffffff")).toBeCloseTo(4.48, 1)
  })

  it("flags a default sheet name only once the sheet has content", () => {
    const empty = new SheetController()
    expect(kinds(empty)).toEqual([])
    const c = setup()
    expect(kinds(c)).toEqual(["sheetName"])
    c.sheets.rename(0, "Sales")
    expect(kinds(c)).toEqual([])
  })

  it("finds a chart without alt text; describing it or marking it decorative clears it", () => {
    const c = setup()
    c.sheets.rename(0, "Sales")
    c.setSelection(selectRange(parseRange("A1:B3")!))
    const chart = c.charts.insert()!
    const [issue] = checkAccessibility(c.wb)
    expect(issue).toMatchObject({ kind: "altText", label: chart.name })
    c.charts.setAltText(chart.id, "Sales rise from 40 to 55", false)
    expect(kinds(c)).toEqual([])
    c.charts.setAltText(chart.id, "", false)
    issueActions(c, checkAccessibility(c.wb)[0], () => undefined)
      .find((a) => a.label === "Mark as decorative")!
      .run()
    expect(c.sheet.charts[0].decorative).toBe(true)
    expect(kinds(c)).toEqual([])
  })

  it("finds hard-to-read text and merged cells, and fixes them", () => {
    const c = setup()
    c.sheets.rename(0, "Sales")
    c.setSelection(selectRange(parseRange("A2")!))
    c.format.apply({ fontColor: "#cccccc" })
    c.setSelection(selectRange(parseRange("C1:D1")!))
    c.format.merge("cells")
    const issues = checkAccessibility(c.wb)
    expect(issues.map((i) => [i.kind, i.label])).toEqual([
      ["contrast", "A2"],
      ["merged", "C1:D1"],
    ])
    for (const issue of issues) issueActions(c, issue, () => undefined)[0].run()
    expect(kinds(c)).toEqual([])
  })

  it("flags link text that doesn't say where it goes", () => {
    const c = setup()
    c.sheets.rename(0, "Sales")
    c.wb.transact(
      [
        { kind: "setCells", sheet: 0, cells: [{ row: 5, col: 0, data: { value: "click here" } }] },
        { kind: "hyperlinks", sheet: 0, links: [{ range: parseRange("A6")!, address: "https://example.com" }] },
      ],
      { label: "x", before: c.selection },
    )
    expect(kinds(c)).toEqual(["linkText"])
  })
})
