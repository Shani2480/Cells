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

const sheetXml = async (bytes: Uint8Array) =>
  new TextDecoder().decode((await readZip(bytes)).find((e) => e.name === "xl/worksheets/sheet1.xml")!.data)

describe("outlines in .xlsx", () => {
  it("keeps row and column groups, collapsed groups, Settings and the outline symbols", async () => {
    const wb = budget()
    const s = wb.sheets[0]
    for (const r of [1, 2, 3]) s.rows.setLevel(r, 1)
    s.rows.setLevel(2, 2)
    s.rows.setHidden(2, true)
    s.rows.setCollapsed(3, true)
    s.cols.setLevel(1, 1)
    s.cols.setLevel(2, 1)
    s.outline = { summaryBelow: false, summaryRight: true, showSymbols: false }
    const bytes = await saveXlsx(wb, { now: new Date(Date.UTC(2026, 0, 1)) })
    const xml = await sheetXml(bytes)
    expect(xml).toMatch(/<sheetPr><tabColor [^>]*\/><outlinePr summaryBelow="0"\/><\/sheetPr>/)
    expect(xml).toMatch(/<sheetView [^>]*showOutlineSymbols="0"/)
    expect(xml).toMatch(/<sheetFormatPr[^>]* outlineLevelRow="2" outlineLevelCol="1"/)
    expect(xml).toMatch(/<row r="3"[^>]* hidden="1"[^>]* outlineLevel="2"/)
    expect(xml).toMatch(/<col min="2" max="3"[^>]* outlineLevel="1"/)
    const back = (await openXlsx(bytes, "Book1.xlsx")).workbook.sheets[0]
    expect([0, 1, 2, 3, 4].map((r) => back.rows.level(r))).toEqual([0, 1, 2, 1, 0])
    expect([back.rows.isHidden(2), back.rows.isCollapsed(3), back.cols.level(1), back.cols.level(2)]).toEqual([
      true,
      true,
      1,
      1,
    ])
    expect(back.outline).toEqual({ summaryBelow: false, summaryRight: true, showSymbols: false })
    // Saved again, the levels aren't doubled by kept attributes.
    const again = await sheetXml(await saveXlsx((await openXlsx(bytes, "Book1.xlsx")).workbook))
    expect(again.match(/outlineLevel="2"/g)).toHaveLength(1)
  })
})

describe("What-If data tables", () => {
  it('writes <f t="dataTable"> on the block\'s first cell and reads every cell back as TABLE', async () => {
    const { Workbook } = await import("@cells/core")
    const wb = new Workbook(["Sheet1"])
    const cells: { row: number; col: number; data: { input?: string; value: number } }[] = [
      { row: 0, col: 0, data: { value: 1 } },
      { row: 1, col: 0, data: { value: 1 } },
      { row: 2, col: 0, data: { input: "=A1*A2", value: 1 } },
      { row: 0, col: 2, data: { input: "=A3", value: 1 } },
      { row: 0, col: 3, data: { value: 2 } },
      { row: 0, col: 4, data: { value: 3 } },
      { row: 1, col: 2, data: { value: 10 } },
      { row: 2, col: 2, data: { value: 100 } },
      { row: 4, col: 0, data: { input: "=A3*2", value: 2 } },
      { row: 5, col: 0, data: { value: 5 } },
      { row: 5, col: 1, data: { input: "=TABLE(,A1)", value: 0 } },
    ]
    for (const [row, col] of [
      [1, 3],
      [1, 4],
      [2, 3],
      [2, 4],
    ])
      cells.push({ row, col, data: { input: "=TABLE(A1,A2)", value: 0 } })
    // A one-input table: the formula above it (B5), the value to try left of it (A6).
    cells.push({ row: 4, col: 1, data: { input: "=A5", value: 2 } })
    wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load", history: false })
    wb.calculate(true)
    const bytes = await saveXlsx(wb)
    const entries = await readZip(bytes)
    const sheet = new TextDecoder().decode(entries.find((e) => e.normalizedName === "xl/worksheets/sheet1.xml")!.data)
    expect(sheet).toContain(`<c r="D2"><f t="dataTable" ref="D2:E3" dt2D="1" dtr="1" r1="A1" r2="A2"/><v>20</v></c>`)
    expect(sheet).toContain(`<c r="E3"><v>300</v></c>`)
    expect(sheet).toContain(`<c r="B6"><f t="dataTable" ref="B6" dt2D="0" dtr="0" r1="A1"/><v>10</v></c>`)
    const back = (await openXlsx(bytes)).workbook
    const s = back.sheets[0]
    expect(s.cells.input(2, 4)).toBe("=TABLE(A1,A2)")
    expect(s.cells.input(5, 1)).toBe("=TABLE(,A1)")
    back.calculate(true)
    expect([s.cells.value(1, 3), s.cells.value(1, 4), s.cells.value(2, 3), s.cells.value(2, 4)]).toEqual([
      20, 30, 200, 300,
    ])
    expect(s.cells.value(5, 1)).toBe(10)
  })
})

describe("Scenarios", () => {
  it("are written as <scenarios> and read back", async () => {
    const { Workbook } = await import("@cells/core")
    const wb = new Workbook(["Sheet1"])
    wb.sheets[0].cells.set(0, 0, { value: 1 })
    wb.sheets[0].scenarios = {
      scenarios: [
        {
          name: "Best & <high>",
          comment: "Created by Alex on 17/11/2026",
          user: "Alex",
          locked: true,
          hidden: false,
          cells: [
            { row: 0, col: 0, value: 20 },
            { row: 1, col: 0, value: "text" },
          ],
        },
        { name: "Hidden", comment: "", user: "", locked: false, hidden: true, cells: [{ row: 0, col: 0, value: 5 }] },
      ],
      current: 0,
      show: 1,
      resultCells: [{ r0: 2, c0: 0, r1: 3, c1: 0 }],
    }
    const bytes = await saveXlsx(wb)
    const entries = await readZip(bytes)
    const sheet = new TextDecoder().decode(entries.find((e) => e.normalizedName === "xl/worksheets/sheet1.xml")!.data)
    expect(sheet).toContain(
      `<scenarios current="0" show="1" sqref="A3:A4"><scenario name="Best &amp; &lt;high&gt;" locked="1" count="2" user="Alex" comment="Created by Alex on 17/11/2026"><inputCells r="A1" val="20"/><inputCells r="A2" val="text"/></scenario>`,
    )
    const back = (await openXlsx(bytes)).workbook
    expect(back.sheets[0].scenarios).toEqual(wb.sheets[0].scenarios)
  })
})
