// Print settings in xlsx: a file Excel 365 saved with every Page Setup option it has (fixtures/excel/print.xlsx), read
// into each sheet's page setup and written back as Excel writes them.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { MAX_COLS, paginate, Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { parseHeaderFooter } from "../src/xlsx/page-setup"
import { readZip } from "../src/zip"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/print.xlsx", import.meta.url)))
const files = async (bytes: Uint8Array) =>
  new Map((await readZip(bytes)).map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))

describe("page setup in xlsx", () => {
  it("reads everything Excel's Page Setup wrote", async () => {
    const { workbook } = await openXlsx(fixture(), "print.xlsx")
    const [report, fit] = workbook.sheets.map((s) => s.pageSetup)
    expect(report).toMatchObject({
      orientation: "landscape",
      paperSize: 9,
      scale: 80,
      fitTo: null,
      centerHorizontally: true,
      centerVertically: false,
      gridlines: true,
      headings: true,
      header: { left: "", center: "&A", right: "" },
      footer: { left: "&D", center: "", right: "Page &P of &N" },
      printArea: [{ r0: 0, c0: 0, r1: 59, c1: 7 }],
      titleRows: [0, 1],
      titleCols: [0, 0],
      rowBreaks: [30],
      colBreaks: [],
    })
    expect(report.margins).toEqual({ left: 0.5, right: 0.5, top: 1, bottom: 0.9, header: 0.4, footer: 0.35 })
    expect(fit).toMatchObject({
      fitTo: { width: 1, height: 0 },
      order: "overThenDown",
      firstPageNumber: 5,
      colBreaks: [10],
      printArea: null,
    })
  })

  it("writes it back as Excel does, the printer settings kept", async () => {
    const read = await openXlsx(fixture(), "print.xlsx")
    const out = await files(await saveXlsx(read.workbook, { source: read.source }))
    const s1 = out.get("xl/worksheets/sheet1.xml")!
    expect(s1).toContain('<printOptions horizontalCentered="1" headings="1" gridLines="1"/>')
    expect(s1).toContain('<pageSetup paperSize="9" scale="80" orientation="landscape" r:id="rId1"/>')
    expect(s1).toContain(
      "<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddFooter>&amp;L&amp;D&amp;RPage &amp;P of &amp;N</oddFooter></headerFooter>",
    )
    expect(s1).toMatch(/<rowBreaks count="1" manualBreakCount="1"><brk id="30" max="\d+" man="1"\/><\/rowBreaks>/)
    const s2 = out.get("xl/worksheets/sheet2.xml")!
    expect(s2).toContain('<pageSetUpPr fitToPage="1"/>')
    expect(s2).toMatch(/<pageSetup [^>]*fitToHeight="0"[^>]*pageOrder="overThenDown"[^>]*useFirstPageNumber="1"/)
    expect(out.get("xl/workbook.xml")).toContain(
      '<definedName name="_xlnm.Print_Area" localSheetId="0">Report!$A$1:$H$60</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Report!$A:$A,Report!$1:$2</definedName>',
    )
    const again = await openXlsx(await saveXlsx(read.workbook, { source: read.source }), "print.xlsx")
    expect(again.workbook.sheets.map((s) => s.pageSetup)).toEqual(read.workbook.sheets.map((s) => s.pageSetup))
  })

  it("a new workbook's settings: margins always, the rest only when set", async () => {
    const wb = new Workbook(["Sheet1"])
    let out = await files(await saveXlsx(wb))
    let xml = out.get("xl/worksheets/sheet1.xml")!
    expect(xml).toContain('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    expect(xml).not.toMatch(/<pageSetup|<printOptions|<headerFooter|<sheetPr/)
    const p = wb.sheets[0].pageSetup
    p.fitTo = { width: 1, height: 1 }
    p.printArea = [
      { r0: 0, c0: 0, r1: 9, c1: 3 },
      { r0: 4, c0: 0, r1: 6, c1: MAX_COLS - 1 },
    ]
    p.titleRows = [0, 0]
    out = await files(await saveXlsx(wb))
    xml = out.get("xl/worksheets/sheet1.xml")!
    expect(xml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
    expect(out.get("xl/workbook.xml")).toContain(
      '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$10,Sheet1!$5:$7</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>',
    )
  })

  it("pages as Excel counted them (PageSetup.Pages: Report 2, Fit 1)", async () => {
    const { workbook } = await openXlsx(fixture(), "print.xlsx")
    const [report, fit] = workbook.sheets.map((s) => paginate(s, s.pageSetup))
    expect(report.pages.map((p) => [p.rows, p.cols, p.titleRows, p.titleCols])).toEqual([
      [[0, 29], [0, 7], null, null],
      [[30, 59], [0, 7], [0, 1], null],
    ])
    expect(fit.pages).toHaveLength(1)
    expect(fit.pages[0].number).toBe(5)
  })

  it("header and footer sections", () => {
    expect(parseHeaderFooter('&L&"Arial,Bold"Left&CMiddle && more&RRight')).toEqual({
      left: '&"Arial,Bold"Left',
      center: "Middle && more",
      right: "Right",
    })
    expect(parseHeaderFooter("Just text")).toEqual({ left: "", center: "Just text", right: "" })
  })
})
