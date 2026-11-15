import { describe, expect, it } from "vitest"
import { copiedRangeInfo, rangeHtml, Workbook } from "../src"

function sheet() {
  const wb = new Workbook(["Sales"])
  const s = wb.sheets[0]
  const bold = wb.styles.intern({ bold: true, fill: "#ffff00", borderBottom: { style: "thin", color: "#ff0000" } })
  s.cells.set(0, 0, { value: "Item", style: bold })
  s.cells.set(0, 1, { value: "Amount", style: bold })
  s.cells.set(1, 0, { value: "Rent <home>" })
  s.cells.set(1, 1, { value: 1200 })
  s.cells.set(2, 0, { value: "Merged" })
  s.merges.add({ r0: 2, c0: 0, r1: 2, c1: 1 })
  return { wb, s }
}

describe("a copied range as HTML", () => {
  it("carries each cell's look, merges and sizes, and where it came from", () => {
    const { wb, s } = sheet()
    const range = { r0: 0, c0: 0, r1: 2, c1: 1 }
    const html = rangeHtml(
      s,
      wb.styles,
      range,
      [
        ["Item", "Amount"],
        ["Rent <home>", "1,200"],
        ["Merged", ""],
      ],
      {
        file: "/home/alex/Q1 Budget.xlsx",
        sheet: "Sales",
        range,
        link: "cells://open?file=%2Fhome%2Falex%2FQ1%20Budget.xlsx&range=Sales!A1%3AB3",
      },
    )
    expect(html).toContain("font-weight:700")
    expect(html).toContain("background:#ffff00")
    expect(html).toContain("border-bottom:1px solid #ff0000")
    expect(html).toContain("Rent &lt;home&gt;")
    // Numbers right-aligned as General shows them.
    expect(html).toMatch(/text-align:right[^>]*>1,200</)
    expect(html).toContain('colspan="2"')
    expect((html.match(/<td/g) ?? []).length).toBe(5)
    expect((html.match(/<col /g) ?? []).length).toBe(2)
    expect(copiedRangeInfo(html)).toEqual({
      app: "cells",
      version: 1,
      file: "/home/alex/Q1 Budget.xlsx",
      sheet: "Sales",
      range: "A1:B3",
      link: "cells://open?file=%2Fhome%2Falex%2FQ1%20Budget.xlsx&range=Sales!A1%3AB3",
    })
  })

  it("leaves hidden rows out, and HTML from elsewhere has no note", () => {
    const { wb, s } = sheet()
    s.rows.setHidden(1, true)
    const html = rangeHtml(s, wb.styles, { r0: 0, c0: 0, r1: 2, c1: 1 }, [
      ["a", "b"],
      ["c", "d"],
      ["e", ""],
    ])
    expect(html).not.toContain(">c<")
    expect(copiedRangeInfo(html)).toBeNull()
    expect(copiedRangeInfo("<table><tr><td>1</td></tr></table>")).toBeNull()
  })
})
