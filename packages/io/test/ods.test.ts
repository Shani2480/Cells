import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { openXlsx } from "../src/xlsx/read"
import { openOds } from "../src/ods/read"
import { saveOds } from "../src/ods/write"
import { formulaFromOds, formulaToOds } from "../src/ods/formula"
import { formatFromOds, formatToOds, NS_NUMBER } from "../src/ods/numfmt"
import { parseTree } from "../src/xlsx/dom"
import { snapshot as full } from "./workbook-snapshot"

/** An .ods has no theme: colours are plain ones, and Excel's built-in cell styles are just formats. */
const snapshot = (...args: Parameters<typeof full>) => {
  const snap = full(...args)
  // A merge's hidden cells: an .ods keeps no format for them.
  for (const s of snap.sheets)
    for (const key of Object.keys(s.cells)) {
      const [r, c] = key.split(",").map(Number)
      const hidden = s.merges.some(
        (m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1 && (r !== m.r0 || c !== m.c0),
      )
      if (hidden) delete s.cells[key]
    }
  return JSON.parse(
    JSON.stringify(snap, (k, v) =>
      k === "fontThemeColor" || k === "fillThemeColor" || k === "cellStyle" ? undefined : v,
    ),
  )
}

const file = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/ods/${name}`, import.meta.url)))

describe(".ods (OpenDocument Spreadsheet)", () => {
  it("Excel's .ods opens as Excel's .xlsx of the same workbook does", async () => {
    const ods = await openOds(file("excel-demo.ods"), "excel-demo.ods")
    const xlsx = await openXlsx(file("demo.xlsx"), "demo.xlsx")
    expect(ods.warnings).toEqual([])
    const a = snapshot(ods.workbook, ods.views)
    const b = snapshot(xlsx.workbook, xlsx.views)
    // Excel doesn't write frozen panes into an .ods.
    for (const s of [...a.sheets, ...b.sheets]) s.view.freeze = null
    expect(a).toEqual(b)
  })

  it("LibreOffice's .ods (made from Excel's .xlsx) opens with the same cells, formulas and layout", async () => {
    const ods = await openOds(file("libreoffice-demo.ods"), "libreoffice-demo.ods")
    const xlsx = await openXlsx(file("demo.xlsx"), "demo.xlsx")
    expect(ods.warnings).toEqual([])
    const a = snapshot(ods.workbook, ods.views)
    const b = snapshot(xlsx.workbook, xlsx.views)
    // LibreOffice measures columns and fitted rows with its own fonts, and a converted file has no view settings.
    const sizes = (s: { widths: Record<string, number>; heights: Record<string, number>; view: unknown }) => {
      const out = { widths: s.widths, heights: s.heights }
      s.widths = {}
      s.heights = {}
      s.view = null
      return out
    }
    const [la, lb] = [a.sheets.map(sizes), b.sheets.map(sizes)]
    expect(a).toEqual(b)
    expect(la[0].widths[9]).toBe(-1)
    expect(Math.abs(la[0].heights[8] - lb[0].heights[8])).toBeLessThanOrEqual(1)
    expect(la[0].widths[0]).toBeGreaterThan(80)
  })

  it("what Cells saves as .ods opens again the same", async () => {
    for (const name of ["demo.xlsx", "../xlsb/more.xlsx"]) {
      const xlsx = await openXlsx(file(name), name)
      const bytes = await saveOds(xlsx.workbook, { views: xlsx.views, activeSheet: xlsx.activeSheet })
      expect(new TextDecoder().decode(bytes.slice(30, 38))).toBe("mimetype")
      const ods = await openOds(bytes, "round.ods")
      const a = snapshot(ods.workbook, ods.views)
      const b = snapshot(xlsx.workbook, xlsx.views)
      // An .ods has no tables: their references become the cells they stand for, with the same results.
      const structured: string[] = []
      for (const [i, s] of b.sheets.entries())
        for (const [key, c] of Object.entries(s.cells) as [string, { input?: string }][])
          if (c.input?.includes("[")) {
            structured.push(`${a.sheets[i].cells[key].input}`)
            delete c.input
            delete a.sheets[i].cells[key].input
          }
      expect(a, name).toEqual(b)
      if (name.includes("more")) expect(structured).toContain("=SUM($D$2:$D$4)")
      if (name.includes("more")) expect(structured).toContain("=$B2*$C2")
    }
  })
})

describe("OpenFormula", () => {
  const ctx = { tables: [], row: 0, sheet: "Sales" }
  it.each([
    ["=SUM(A1:B2,Rates!C3)", "of:=SUM([.A1:.B2];[$Rates.C3])"],
    ["=SUM(Sales:Notes!H1)", "of:=SUM([$Sales.H1:$Notes.H1])"],
    ['=COUNTIF(H:H,">2")+SUM(1:1)', 'of:=COUNTIF([.H$1:.H$1048576];">2")+SUM([.$A1:.$XFD1])'],
    ["=SUM((A1:A2,B1:B2))", "of:=SUM(([.A1:.A2]~[.B1:.B2]))"],
    ["=A1:B3 B2:C4", "of:=[.A1:.B3]![.B2:.C4]"],
    ["=LET(x,2,x*x)", "of:=COM.MICROSOFT.LET(_xlpm.x;2;_xlpm.x*_xlpm.x)"],
    ['=INDEX({"x","y";TRUE,#N/A},2,1)', 'of:=INDEX({"x";"y"|TRUE;#N/A};2;1)'],
    ["=IF(TRUE,1,2)", "of:=IF(TRUE;1;2)"],
    ["=XLOOKUP(1,A:A,B:B)", "of:=COM.MICROSOFT.XLOOKUP(1;[.A$1:.A$1048576];[.B$1:.B$1048576])"],
    ["=CHIDIST(1,2)+F.DIST(1,2,3,TRUE)", "of:=LEGACY.CHIDIST(1;2)+FDIST(1;2;3;TRUE)"],
  ])("%s ⇄ %s", (excel, ods) => {
    expect(formulaToOds(excel, ctx)).toBe(ods)
    expect("=" + formulaFromOds(ods)).toBe(excel)
  })

  it("reads LibreOffice's whole columns and rows, and leaves references to other files out", () => {
    expect(formulaFromOds("of:=SUM([.H:.H])+SUM([.1:.1])")).toBe("SUM(H:H)+SUM(1:1)")
    expect(formulaFromOds("of:=IF(TRUE();1;FALSE())")).toBe("IF(TRUE,1,FALSE)")
    expect(formulaFromOds("of:=['file:///home/alex/other.ods'#$Sheet1.A1]")).toBeNull()
    expect(formulaToOds("=A1#", ctx)).toBeNull()
  })
})

describe("OpenDocument number formats", () => {
  const roundTrip = (code: string) => {
    const { xml, type } = formatToOds(code, "N1")
    const doc = parseTree(
      new TextEncoder().encode(
        `<r xmlns:number="${NS_NUMBER}" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">${xml}</r>`,
      ),
    )
    const byName = new Map(
      doc.root.children.map((c) => [c.nsAttrs[`urn:oasis:names:tc:opendocument:xmlns:style:1.0|name`] ?? "", c]),
    )
    return { code: formatFromOds(byName.get("N1")!, (n) => byName.get(n)), type }
  }
  it.each([
    ["0.00", "float"],
    ["#,##0", "float"],
    ["0.0%", "percentage"],
    ["yyyy-mm-dd", "date"],
    ["[h]:mm:ss", "time"],
    ["#,##0.00;[Red]-#,##0.00", "float"],
    ["@", "string"],
  ])("%s", (code, type) => {
    expect(roundTrip(code)).toEqual({ code, type })
  })
})
