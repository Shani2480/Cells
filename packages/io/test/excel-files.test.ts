// Files Excel itself wrote (the made-up samples, opened and saved again by Excel 365): Cells reads them as it wrote
// them, and saving keeps the parts it doesn't read.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { passwordMatches, type CellStyle, type Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"
import { samples } from "./samples"

const fixture = (name: string) =>
  new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}.xlsx`, import.meta.url)))

/** Excel writes number formats in its own spelling: "-" and " " escaped. Both show the same. */
const canonical = (style: CellStyle) =>
  style.numberFormat ? { ...style, numberFormat: style.numberFormat.replace(/\\([- ])/g, "$1") } : style

/** Every stored cell: value, formula and format, by sheet. */
function dump(wb: Workbook) {
  return wb.sheets.map((s) => {
    const cells: unknown[] = []
    const used = s.cells.usedBounds()
    if (used)
      s.cells.forEach(used, (row, col, d) =>
        cells.push([row, col, d.value, d.spill ? undefined : d.input, canonical(wb.styles.get(s.styleAt(row, col)))]),
      )
    return { name: s.name, hidden: s.hidden, cells }
  })
}

describe.each(["budget", "formats", "conditional", "layout", "protected"])("Excel's %s.xlsx", (name) => {
  it("reads as the workbook Cells saved", async () => {
    const read = await openXlsx(fixture(name), `${name}.xlsx`)
    const original = samples[name]()
    expect(dump(read.workbook)).toEqual(dump(original))
    read.workbook.sheets.forEach((s, i) => {
      const o = original.sheets[i]
      expect(s.merges.all).toEqual(o.merges.all)
      expect(s.conditional).toEqual(o.conditional)
      expect(s.annotations.all()).toEqual(o.annotations.all())
      expect(s.protection).toEqual(o.protection)
      expect(s.tabColor).toEqual(o.tabColor)
      for (let c = 0; c < 20; c++) expect([c, s.cols.size(c)]).toEqual([c, o.cols.size(c)])
      // Rows Excel fitted to their fonts (no customHeight) are left out: Cells fits them in the app.
      for (let r = 0; r < 20; r++) if (!s.rows.isAuto(r)) expect([r, s.rows.size(r)]).toEqual([r, o.rows.size(r)])
    })
    expect(read.workbook.names.all()).toEqual(original.names.all())
    expect(read.workbook.structure.protection).toEqual(original.structure.protection)
  })

  it("saves back with Excel's own parts kept", async () => {
    const bytes = fixture(name)
    const read = await openXlsx(bytes, `${name}.xlsx`)
    const saved = await saveXlsx(read.workbook, {
      source: read.source,
      views: read.views,
      activeSheet: read.activeSheet,
    })
    const before = new Map((await readZip(bytes)).map((e) => [e.normalizedName, e.data]))
    const after = new Map((await readZip(saved)).map((e) => [e.normalizedName, e.data]))
    for (const kept of ["xl/theme/theme1.xml", "docProps/core.xml", "docProps/app.xml"])
      expect(after.get(kept), kept).toEqual(before.get(kept))
    expect(after.has("xl/calcChain.xml")).toBe(false)
    const again = await openXlsx(saved, `${name}.xlsx`)
    expect(dump(again.workbook)).toEqual(dump(read.workbook))
  })
})

describe("passwords Excel 365 saved (SHA-512, 100,000 spins)", () => {
  it("unlock with the right password only, and stay as they were when saved", async () => {
    const read = await openXlsx(fixture("passwords"), "passwords.xlsx")
    const sheet = read.workbook.sheets[0].protection!
    const book = read.workbook.structure.protection!
    expect(sheet.hash?.spinCount).toBe(100_000)
    const t = performance.now()
    expect(passwordMatches(sheet, "Secret 1!")).toBe(true)
    const ms = performance.now() - t
    expect(passwordMatches(sheet, "secret 1!")).toBe(false)
    expect(passwordMatches(book, "Book pw")).toBe(true)
    expect(passwordMatches(book, "Secret 1!")).toBe(false)
    expect(ms).toBeLessThan(1500)
    const saved = await openXlsx(await saveXlsx(read.workbook, { source: read.source }), "passwords.xlsx")
    expect(saved.workbook.sheets[0].protection).toEqual(sheet)
    expect(saved.workbook.structure.protection).toEqual(book)
  })
})

describe("Data Validation Excel 365 saved", () => {
  const rules = (wb: Workbook) =>
    wb.sheets[0].validations.map((e) => ({ ranges: e.ranges.map((r) => [r.r0, r.c0, r.r1, r.c1]), rule: e.rule }))

  it("reads every kind, with the other-sheet list from the x14 extension", async () => {
    const read = await openXlsx(fixture("validation"), "validation.xlsx")
    const list = rules(read.workbook)
    expect(list.map((e) => e.rule.type)).toEqual([
      "whole",
      "decimal",
      "list",
      "date",
      "textLength",
      "time",
      "custom",
      "list",
    ])
    const qty = list[0].rule
    expect(qty).toMatchObject({ operator: "between", formula1: "1", formula2: "100", allowBlank: true })
    expect(qty).toMatchObject({ promptTitle: "Quantity", prompt: "Whole number from 1 to 100", showInput: true })
    expect(qty).toMatchObject({ errorTitle: "Bad quantity", error: "Type 1 to 100.", errorStyle: "stop" })
    expect(list[1].rule).toMatchObject({ operator: "greaterThanOrEqual", formula1: "0", errorStyle: "warning" })
    expect(list[2].rule).toMatchObject({ formula1: '"Apple,Banana,Cherry"', dropDown: true })
    expect(list[3].rule).toMatchObject({ formula1: "DATE(2026,1,1)", formula2: "DATE(2026,12,31)" })
    expect(list[4].rule).toMatchObject({ operator: "lessThanOrEqual", formula1: "5", errorStyle: "information" })
    expect(list[6]).toMatchObject({
      ranges: [
        [1, 7, 9, 7],
        [1, 9, 9, 9],
      ],
      rule: { formula1: "MOD(H2,2)=0" },
    })
    expect(list[7]).toMatchObject({
      ranges: [[1, 5, 19, 5]],
      rule: { formula1: "Lists!$A$1:$A$3", errorStyle: "warning" },
    })
  })

  it("saves the rules back where Excel keeps them", async () => {
    const read = await openXlsx(fixture("validation"), "validation.xlsx")
    const saved = await saveXlsx(read.workbook, {
      source: read.source,
      views: read.views,
      activeSheet: read.activeSheet,
    })
    const files = new Map((await readZip(saved)).map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
    const xml = files.get("xl/worksheets/sheet1.xml")!
    expect(xml).toContain('<dataValidations count="7">')
    expect(xml).toMatch(
      /<x14:dataValidations count="1"[^>]*>.*<xm:f>Lists!\$A\$1:\$A\$3<\/xm:f>.*<xm:sqref>F2:F20<\/xm:sqref>/,
    )
    expect(xml.match(/CCE6A557/g)).toHaveLength(1)
    const again = await openXlsx(saved, "validation.xlsx")
    expect(rules(again.workbook)).toEqual(rules(read.workbook))
  })
})

describe("tables", () => {
  it("reads Excel's table with its total row, and shows its references as Excel does", async () => {
    const read = await openXlsx(fixture("features"), "features.xlsx")
    const sheet = read.workbook.sheets[0]
    expect(sheet.tables).toHaveLength(1)
    const t = sheet.tables[0]
    expect(t).toMatchObject({ name: "SalesTable", range: { r0: 0, c0: 0, r1: 6, c1: 3 }, header: true, totals: true })
    expect(t.columns.map((c) => c.name)).toEqual(["Region", "Q1", "Q2", "Total"])
    expect(t.columns[0].totalsLabel).toBe("Total")
    expect(t.columns[3]).toMatchObject({ totals: "sum", formula: "B2+C2" })
    expect(t).toMatchObject({ style: "TableStyleMedium2", rowStripes: true, columnStripes: false })
    expect(sheet.cells.get(6, 3)?.input).toBe("=SUBTOTAL(109,[Total])")
    expect(sheet.cells.get(1, 5)?.input).toBe("=SUM(SalesTable[Total])")
    expect(sheet.cells.get(1, 5)?.value).toBe(sheet.cells.get(6, 3)?.value)
  })

  it("writes tables back as Excel does, with this-row references spelled out", async () => {
    const wb = samples.tables()
    expect(wb.sheets[0].cells.get(1, 3)?.value).toBe(3)
    expect(wb.sheets[1].cells.get(0, 0)?.value).toBe(52)
    const saved = await saveXlsx(wb)
    const files = new Map((await readZip(saved)).map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
    const table = files.get("xl/tables/table1.xml")!
    expect(table).toContain('ref="A1:D6" totalsRowCount="1"')
    expect(table).toContain("<calculatedColumnFormula>Orders[[#This Row],[Qty]]*Orders[[#This Row],[Unit Price]]")
    expect(table).toContain('<tableStyleInfo name="TableStyleLight9" showFirstColumn="1"')
    expect(files.get("xl/worksheets/sheet1.xml")).toContain(
      "<f>Orders[[#This Row],[Qty]]*Orders[[#This Row],[Unit Price]]</f>",
    )
    expect(files.get("xl/worksheets/sheet1.xml")).toMatch(
      /<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts>/,
    )
    expect(files.get("[Content_Types].xml")).toContain("/xl/tables/table1.xml")
    const again = await openXlsx(saved, "tables.xlsx")
    expect(again.workbook.sheets[0].tables).toEqual(wb.sheets[0].tables)
    expect(dump(again.workbook)).toEqual(dump(wb))
    // Converted to a range: the table part goes.
    again.workbook.sheets[0].tables = []
    const plain = await saveXlsx(again.workbook, { source: again.source })
    const names = (await readZip(plain)).map((e) => e.normalizedName)
    expect(names.some((n) => n.startsWith("xl/tables/"))).toBe(false)
  })
})
