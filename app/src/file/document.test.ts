import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { classifyOpenError, openDocument } from "./document"

const corpus = () => new Uint8Array(readFileSync(join(__dirname, "../../../packages/io/test/fixtures/xls/corpus.xls")))

describe("an Excel 97-2003 workbook (.xls)", () => {
  it("opens as an unsaved .xlsx copy, saying so first", async () => {
    const doc = await openDocument(corpus(), { name: "Sales.xls", path: "/home/alex/Sales.xls" }, undefined, false)
    expect(doc.ref).toBeNull()
    expect(doc.fileName).toBe("Sales.xlsx")
    expect(doc.warnings[0]).toMatch(/Excel 97-2003 workbook.*\.xlsx/)
    const s = doc.controller.wb.sheets[0]
    expect(s.cells.input(7, 1)).toBe("=EOMONTH(A8,1)")
    expect(s.cells.value(5, 3)).toBe("big!")
    const bytes = await doc.bytes()
    expect([bytes[0], bytes[1]]).toEqual([0x50, 0x4b])
    doc.close()
  })

  it("explains one it can't read", async () => {
    const broken = corpus().slice(0, 1024)
    const error = await openDocument(broken, { name: "Old.xls", path: "/home/alex/Old.xls" }, undefined, false).catch(
      (e: unknown) => e,
    )
    expect(classifyOpenError(error, broken, "Old.xls").kind).toBe("damaged")
  })
})

const fixture = (path: string) =>
  new Uint8Array(readFileSync(join(__dirname, "../../../packages/io/test/fixtures/", path)))

describe("an Excel Binary Workbook (.xlsb)", () => {
  it("opens as an unsaved .xlsx copy, saying so first", async () => {
    const doc = await openDocument(
      fixture("xlsb/demo.xlsb"),
      { name: "Sales.xlsb", path: "/home/alex/Sales.xlsb" },
      undefined,
      false,
    )
    expect(doc.ref).toBeNull()
    expect(doc.fileName).toBe("Sales.xlsx")
    expect(doc.warnings[0]).toMatch(/Excel Binary Workbook.*\.xlsx/)
    expect(doc.controller.wb.sheets.map((s) => s.name)).toEqual(["Sales", "Rates"])
    const bytes = await doc.bytes()
    expect([bytes[0], bytes[1]]).toEqual([0x50, 0x4b])
    doc.close()
  })

  it("explains one it can't read", async () => {
    const broken = fixture("xlsb/demo.xlsb").slice(0, 400)
    const error = await openDocument(broken, { name: "Old.xlsb", path: "/home/alex/Old.xlsb" }, undefined, false).catch(
      (e: unknown) => e,
    )
    expect(classifyOpenError(error, broken, "Old.xlsb").kind).toBe("damaged")
  })
})

describe("an OpenDocument Spreadsheet (.ods)", () => {
  it("opens as itself, and saves back as .ods", async () => {
    const doc = await openDocument(
      fixture("ods/libreoffice-demo.ods"),
      { name: "Sales.ods", path: "/home/alex/Sales.ods" },
      undefined,
      false,
    )
    expect(doc.ref?.path).toBe("/home/alex/Sales.ods")
    expect(doc.fileName).toBe("Sales.ods")
    expect(doc.warnings).toEqual([])
    const bytes = await doc.bytes()
    expect(new TextDecoder().decode(bytes.slice(30, 38))).toBe("mimetype")
    doc.close()
  })
})

describe("unsaved changes", () => {
  it("adding a custom view makes the workbook unsaved (it isn't an undo step)", async () => {
    const bytes = new Uint8Array(
      readFileSync(join(__dirname, "../../../packages/io/test/fixtures/excel/custom-views.xlsx")),
    )
    const doc = await openDocument(bytes, { name: "Views.xlsx", path: "/home/alex/Views.xlsx" }, undefined, false)
    expect(doc.dirty).toBe(false)
    doc.controller.customViews.add("Mine", true, true)
    expect(doc.dirty).toBe(true)
    doc.close()
  })
})
