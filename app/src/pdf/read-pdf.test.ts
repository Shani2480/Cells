import { readFileSync } from "node:fs"
import { join } from "node:path"
import { pdfFromImages, pdfTables } from "@cells/io"
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs"
import { describe, expect, it } from "vitest"
import { readPdf, type PdfJs } from "./read-pdf"

/** Demo PDFs: ruled.pdf and statement.pdf made by Excel 365, invoice.pdf and rules.pdf by LibreOffice. */
const fixture = (name: string) =>
  new Uint8Array(readFileSync(join(__dirname, "../../../packages/io/test/fixtures/pdf", name)))

const items = async (name: string) => pdfTables((await readPdf(fixture(name), pdfjs as unknown as PdfJs)).pages)

describe("From PDF: tables found in real PDFs", () => {
  it("takes a bordered table's cells from its lines, leaving the title and notes around it out", async () => {
    const list = await items("ruled.pdf")
    expect(list.map((i) => [i.id, i.name, i.kind])).toEqual([
      ["Table001", "Table001 (Page 1)", "Table"],
      ["Page001", "Page001", "Page"],
    ])
    const t = list[0].rows
    // As Excel's own Pdf.Tables finds it with EnforceBorderLines.
    expect(t).toHaveLength(14)
    expect(t[0]).toEqual(["Product", "Region", "Units", "Unit price", "Order date"])
    expect(t[1]).toEqual(["Pens", "North", "10", "1.50", "2026-01-03"])
    expect(t[13]).toEqual(["Total", "", "582", "", ""])
    // The page: every line, the title in the first column.
    expect(list[1].rows[0][0]).toBe("Northwind Supplies - Quarterly sales")
    expect(list[1].rows.at(-1)?.[0]).toBe("Thank you for your business.")
  })

  it("splits columns without lines where the text lines up, and joins a table carried over three pages", async () => {
    const list = await items("statement.pdf")
    expect(list.map((i) => i.name)).toEqual(["Table001 (Page 1-3)", "Page001", "Page002", "Page003"])
    const t = list[0]
    expect(t.pages).toEqual([1, 2, 3])
    // The heading once (it's repeated on every page), then all 90 lines of the statement.
    expect(t.rows).toHaveLength(91)
    expect(t.rows[0]).toEqual(["Date", "Description", "Amount", "Balance"])
    expect(t.rows[1]).toEqual(["01/01/2026", "Coffee shop", "-5.00", "995.00"])
    expect(t.rows[3]).toEqual(["03/01/2026", "Salary", "2,500.00", "3,453.00"])
    expect(t.rows[90]).toEqual(["31/03/2026", "Gym membership", "-58.00", "18,376.00"])
    expect(t.rows.filter((r) => r[0] === "Date")).toHaveLength(1)
  })

  it("keeps a description wrapped over lines in its row (set in the middle of it), its lines apart as Excel's", async () => {
    const [t] = await items("invoice.pdf")
    expect(t.rows).toEqual([
      ["Item", "Description", "Qty", "Price", "Total"],
      [
        "A-100",
        "Logo design, three concepts and two rounds of changes to the chosen\none",
        "1",
        "€ 450.00",
        "€ 450.00",
      ],
      ["A-210", "Business cards", "500", "€ 0.12", "€ 60.00"],
      ["B-007", "Website banner in four sizes for the spring campaign, with source files", "4", "€ 35.00", "€ 140.00"],
      ["C-300", "Printing", "2", "€ 18.50", "€ 37.00"],
    ])
  })

  it("finds a table ruled only across, and a bordered one below it", async () => {
    const list = await items("rules.pdf")
    expect(list.map((i) => i.name)).toEqual(["Table001 (Page 1)", "Table002 (Page 1)", "Page001"])
    expect(list[0].rows).toEqual([
      ["City", "Population", "Area (km²)", "Density"],
      ["Sampleton", "120 500", "85.2", "1 414"],
      ["Demoville", "98 020", "61.0", "1 607"],
      ["Exampleburg", "45 300", "40.7", "1 113"],
      ["Testford", "12 800", "9.9", "1 293"],
    ])
    expect(list[1].rows).toEqual([
      ["Month", "Visits"],
      ["Jan", "310"],
      ["Feb", "290"],
      ["Mar", "355"],
    ])
  })

  it("marks a page that is only a picture as a scan", async () => {
    const { pages } = await readPdf(fixture("ruled.pdf"), pdfjs as unknown as PdfJs)
    expect(pages[0].scanned).toBe(false)
    const jpeg = fixture("scan-statement.jpg")
    const scan = pdfFromImages([{ width: 618, height: 797, jpeg, pixelWidth: 2574, pixelHeight: 3320 }])
    const read = await readPdf(scan, pdfjs as unknown as PdfJs)
    expect(read.pages[0].scanned).toBe(true)
    expect(pdfTables(read.pages).map((i) => [i.name, i.rows.length])).toEqual([["Page001", 0]])
  })
})
