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

describe("text links", () => {
  it("reads Excel's connections and query tables as Cells wrote them, and keeps them on a save", async () => {
    const read = await openXlsx(fixture("linkedText"), "linkedText.xlsx")
    const original = samples.linkedText()
    expect(read.workbook.sheets.map((s) => s.queryTables)).toEqual(original.sheets.map((s) => s.queryTables))
    expect(read.source.keptConnections).toEqual([])
    const again = await openXlsx(await saveXlsx(read.workbook, { source: read.source }))
    expect(again.workbook.sheets.map((s) => s.queryTables)).toEqual(original.sheets.map((s) => s.queryTables))
  })
})

describe("links to other workbooks", () => {
  it("reads Excel's links: formulas by file name, values calculated from the kept copy", async () => {
    const read = await openXlsx(fixture("linked"), "Order.xlsx")
    const wb = read.workbook
    const s = wb.sheets[0]
    expect([0, 1, 2, 3].map((r) => s.cells.input(r, 0))).toEqual([
      "='[Price List.xlsx]Prices'!$B$1*2",
      "=SUM('[Price List.xlsx]Prices'!$B$1:$B$3)",
      "='Price List.xlsx'!Rate+1",
      "=VLOOKUP(\"Ink\",'[Price List.xlsx]Prices'!$A$1:$B$3,2,FALSE)",
    ])
    expect(wb.links.list[0]).toMatchObject({
      name: "Price List.xlsx",
      target: "Price List.xlsx",
      absolute: "C:\\cells-check\\Price List.xlsx",
      names: [{ name: "Rate", refersTo: "='Prices'!$B$2" }],
    })
    wb.calc.calculateNow(true)
    expect([0, 1, 2, 3].map((r) => s.cells.value(r, 0))).toEqual([3, 12.75, 5, 4])
  })

  it("writes them back as Excel does", async () => {
    const read = await openXlsx(fixture("linked"), "Order.xlsx")
    const bytes = await saveXlsx(read.workbook, { source: read.source })
    const entries = await readZip(bytes)
    const text = (name: string) => new TextDecoder().decode(entries.find((e) => e.normalizedName === name)?.data)
    expect(text("xl/worksheets/sheet1.xml")).toContain("<f>[1]Prices!$B$1*2</f>")
    expect(text("xl/worksheets/sheet1.xml")).toContain("<f>[1]!Rate+1</f>")
    expect(text("xl/workbook.xml")).toMatch(
      /<externalReferences><externalReference r:id="rId\d+"\/><\/externalReferences>/,
    )
    expect(text("xl/externalLinks/externalLink1.xml")).toContain('<cell r="B3"><v>7.25</v></cell>')
    expect(text("xl/externalLinks/_rels/externalLink1.xml.rels")).toContain(
      'Target="Price%20List.xlsx" TargetMode="External"',
    )
    expect(entries.filter((e) => e.normalizedName.startsWith("xl/externalLinks/externalLink"))).toHaveLength(1)
    const again = (await openXlsx(bytes)).workbook
    expect(again.links.list).toEqual(read.workbook.links.list)
    expect(again.sheets[0].cells.input(2, 0)).toBe("='Price List.xlsx'!Rate+1")
  })
})

describe("Excel's hyperlinks.xlsx", () => {
  it("reads web, mail, file and place links, and saves them as Excel wrote them", async () => {
    const read = await openXlsx(fixture("hyperlinks"), "hyperlinks.xlsx")
    const links = read.workbook.sheets[0].hyperlinks
    expect(links).toEqual([
      { range: { r0: 0, c0: 0, r1: 0, c1: 0 }, address: "https://example.com/page?a=1&b=2", tooltip: "Go to the site" },
      { range: { r0: 1, c0: 0, r1: 1, c1: 0 }, address: "mailto:alex@example.com?subject=Hello%20there" },
      { range: { r0: 2, c0: 0, r1: 2, c1: 0 }, location: "Data!B2", display: "Totals" },
      // Excel keeps a file's path relative to the workbook's folder.
      { range: { r0: 3, c0: 0, r1: 3, c1: 0 }, address: "../Reports/Q1.pdf" },
      { range: { r0: 0, c0: 1, r1: 1, c1: 2 }, address: "https://example.com/", display: "Block" },
    ])
    const again = await openXlsx(await saveXlsx(read.workbook, { source: read.source }))
    expect(again.workbook.sheets[0].hyperlinks).toEqual(links)
  })
})

describe("Excel's custom-views.xlsx (View › Custom Views)", () => {
  it("reads each view: what it includes, each sheet's zoom and selection, hidden rows and columns, print settings", async () => {
    const read = await openXlsx(fixture("custom-views"), "custom-views.xlsx")
    const wb = read.workbook
    const [noPrint, short, all] = wb.customViews
    expect(wb.customViews.map((v) => v.name)).toEqual(["No print", "Short", "All rows"])
    expect([noPrint.printSettings, noPrint.hiddenRowsCols]).toEqual([false, false])
    const data = wb.sheets[0]
    expect(all.activeSheet).toBe(data)
    const s = short.sheets.get(data)!
    expect(s.zoom).toBe(0.85)
    expect(s.active).toEqual({ row: 11, col: 0 })
    expect(s.hiddenRows).toEqual([[4, 8]])
    expect(s.hiddenCols).toEqual([[1, 1]])
    expect(s.pageSetup?.orientation).toBe("landscape")
    expect(all.sheets.get(data)!.hiddenRows).toEqual([])
    expect(all.sheets.get(data)!.active).toEqual({ row: 4, col: 2 })
    // The view's names are modelled, not kept as read.
    expect(read.source.keptNames.some((n) => n.xml.includes(".wvu."))).toBe(false)
  })

  it("saves them as Excel does, and reads its own file back the same", async () => {
    const read = await openXlsx(fixture("custom-views"), "custom-views.xlsx")
    const bytes = await saveXlsx(read.workbook, { source: read.source, views: read.views })
    const zip = new Map((await readZip(bytes)).map((e) => [e.normalizedName, e.data]))
    const text = (name: string) => new TextDecoder().decode(zip.get(name)!)
    expect(text("xl/workbook.xml")).toContain(
      '<customWorkbookView name="Short" guid="{F947688B-2F8F-43AC-AF3E-986133E91938}"',
    )
    expect(text("xl/workbook.xml")).toContain(
      '<definedName name="Z_F947688B_2F8F_43AC_AF3E_986133E91938_.wvu.Rows" localSheetId="0" hidden="1">Data!$5:$9</definedName>',
    )
    expect(text("xl/worksheets/sheet1.xml")).toMatch(
      /<customSheetView guid="\{F947688B[^"]*\}" scale="85"[^>]*hiddenRows="1"/,
    )
    const again = await openXlsx(bytes, "again.xlsx")
    const strip = (wb: typeof read.workbook) =>
      wb.customViews.map((v) => ({
        ...v,
        activeSheet: wb.sheets.indexOf(v.activeSheet!),
        sheets: [...v.sheets.values()],
      }))
    expect(strip(again.workbook)).toEqual(strip(read.workbook))
  })
})

describe("Excel's edit-ranges.xlsx (Review › Allow Edit Ranges)", () => {
  it("reads each range, its cells and its password, and writes them back the same", async () => {
    const read = await openXlsx(fixture("edit-ranges"), "edit-ranges.xlsx")
    const [inputs, rates] = read.workbook.sheets[0].editRanges
    expect(inputs).toEqual({ name: "Inputs", ranges: [{ r0: 1, c0: 1, r1: 4, c1: 1 }] })
    expect(rates.name).toBe("Rates")
    expect(rates.ranges).toHaveLength(2)
    expect(passwordMatches(rates, "pw")).toBe(true)
    expect(passwordMatches(rates, "nope")).toBe(false)
    const bytes = await saveXlsx(read.workbook, { source: read.source, views: read.views })
    const xml = new TextDecoder().decode(
      (await readZip(bytes)).find((e) => e.normalizedName === "xl/worksheets/sheet1.xml")!.data,
    )
    expect(xml.match(/<protectedRanges>/g)).toHaveLength(1)
    const again = await openXlsx(bytes, "again.xlsx")
    expect(again.workbook.sheets[0].editRanges).toEqual(read.workbook.sheets[0].editRanges)
  })
})

describe("Excel's alt-text.xlsx (charts' Alt Text)", () => {
  it("reads a chart's description and Mark as decorative, and writes them back", async () => {
    const read = await openXlsx(fixture("alt-text"), "alt-text.xlsx")
    const [described, decorative] = read.workbook.sheets[0].charts
    expect(described.altText).toBe("Sales by month, rising")
    expect(described.decorative).toBeUndefined()
    expect(decorative.decorative).toBe(true)
    const again = await openXlsx(await saveXlsx(read.workbook, { source: read.source, views: read.views }), "x.xlsx")
    expect(again.workbook.sheets[0].charts.map((c) => [c.altText, c.decorative])).toEqual([
      ["Sales by month, rising", undefined],
      [undefined, true],
    ])
  })
})
