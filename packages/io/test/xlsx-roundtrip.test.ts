import { describe, expect, it } from "vitest"
import { errorValue } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"
import { budget } from "./samples"

const sample = budget

describe("xlsx round trip", () => {
  it("keeps values, formulas, formats and layout", async () => {
    const wb = sample()
    const bytes = await saveXlsx(wb, { now: new Date(Date.UTC(2026, 0, 1)) })
    const back = await openXlsx(bytes, "Book1.xlsx")
    const b = back.workbook
    expect(b.sheets.map((s) => s.name)).toEqual(["Budget", "Notes"])
    const s = b.sheets[0]
    expect(s.cells.value(0, 0)).toBe("Item")
    expect(s.cells.value(1, 1)).toBe(1200.5)
    expect(s.cells.value(2, 0)).toBe("Food & drink <weekly>")
    expect(s.cells.input(3, 1)).toBe("=SUM(B2:B3)")
    expect(s.cells.value(3, 1)).toBe(1510.5)
    expect(s.cells.value(4, 0)).toBe(true)
    expect(s.cells.value(4, 1)).toEqual(errorValue("#N/A"))
    expect(s.cells.value(5, 0)).toBe("line one\nline two")
    expect(s.cells.input(6, 0)).toBe('=XLOOKUP("Rent",A2:A3,B2:B3)')
    expect(s.cells.value(6, 0)).toBe(1200.5)
    expect(s.cells.input(7, 0)).toBe("=SEQUENCE(3)")
    expect(s.cells.value(9, 0)).toBe(3)
    expect(b.styles.get(s.cells.style(0, 0))).toEqual({ bold: true, fill: "#FFEB9C", numberFormat: "#,##0.00" })
    expect(s.cols.size(0)).toBe(120)
    expect(s.rows.size(5)).toBe(40)
    expect(s.rows.isHidden(10)).toBe(true)
    expect(s.merges.all).toEqual([{ r0: 12, c0: 0, r1: 12, c1: 2 }])
    expect(s.tabColor).toBe("#C00000")
    expect(s.annotations.get(1, 1)).toEqual({ kind: "note", author: "Alex", text: "Paid on the 1st", shown: false })
    expect(s.annotations.get(2, 1)).toEqual(wb.sheets[0].annotations.get(2, 1))
    expect(s.conditional).toEqual(wb.sheets[0].conditional)
    expect(b.names.all()).toEqual([{ name: "Costs", ref: "Budget!$B$2:$B$3" }])
    expect(b.sheets[1].cells.value(0, 0)).toBe(3021)
    expect(back.warnings).toEqual([])
  })

  it("writes the prefixes and dynamic-array markup Excel expects", async () => {
    const bytes = await saveXlsx(sample())
    const entries = await readZip(bytes)
    const text = (name: string) => new TextDecoder().decode(entries.find((e) => e.normalizedName === name)!.data)
    const sheet = text("xl/worksheets/sheet1.xml")
    expect(sheet).toContain("<f>_xlfn.XLOOKUP(")
    expect(sheet).toMatch(/<c r="A8" cm="1"><f t="array" ref="A8:A10">_xlfn.SEQUENCE\(3\)<\/f><v>1<\/v><\/c>/)
    expect(entries.map((e) => e.normalizedName)).toContain("xl/metadata.xml")
    expect(text("[Content_Types].xml")).toContain("/xl/metadata.xml")
  })

  it("saving twice gives the same bytes", async () => {
    const wb = sample()
    const now = new Date(Date.UTC(2026, 0, 1))
    const a = await saveXlsx(wb, { now })
    const opened = await openXlsx(a)
    const b = await saveXlsx(opened.workbook, { source: opened.source, views: opened.views, now })
    const again = await openXlsx(b)
    const c = await saveXlsx(again.workbook, { source: again.source, views: again.views, now })
    expect(c).toEqual(b)
  })
})

describe("after opening", () => {
  it("formulas keep the file's values and recalculate when what they read changes", async () => {
    const opened = await openXlsx(await saveXlsx(budget()))
    const wb = opened.workbook
    expect(wb.sheets[0].cells.value(3, 1)).toBe(1510.5)
    wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 1, col: 1, data: { value: 1000 } }] }], { label: "Edit" })
    expect(wb.sheets[0].cells.value(3, 1)).toBe(1310)
    expect(wb.sheets[1].cells.value(0, 0)).toBe(2620)
  })
})
