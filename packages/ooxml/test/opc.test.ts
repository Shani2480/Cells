import { describe, expect, it } from "vitest"
import {
  allowsMacros,
  CONTENT_TYPE_DOCUMENT,
  CONTENT_TYPE_WORKBOOK,
  decodeXml,
  escapeXml,
  isPassthroughPart,
  mainContentTypeFor,
  OpcError,
  openPackage,
  packageExtensions,
  parseElements,
  readZip,
  SPREADSHEET_PACKAGE,
  storedZip,
  WORD_PACKAGE,
  writePackage,
  writeZip,
  xmlBytes,
  type PackageFormat,
} from "../src"

const text = (s: string) => new TextEncoder().encode(s)

/** The smallest package of a format: content types, package relationships and the main part (plus extras). */
async function minimal(main: string, mainType: string, extra: { name: string; type: string; body: string }[] = []) {
  const overrides = [{ name: main, type: mainType }, ...extra]
    .map((p) => `<Override PartName="/${p.name}" ContentType="${p.type}"/>`)
    .join("")
  return writeZip([
    {
      name: "[Content_Types].xml",
      data: xmlBytes(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
      ),
    },
    {
      name: "_rels/.rels",
      data: xmlBytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${main}"/></Relationships>`,
      ),
    },
    { name: main, data: xmlBytes("<root/>") },
    ...extra.map((p) => ({ name: p.name, data: xmlBytes(p.body) })),
  ])
}

const THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
const STYLES_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
const FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

describe("the OPC layer, per format", () => {
  it("opens and writes a Word document in Word's part order", async () => {
    const bytes = await minimal("word/document.xml", CONTENT_TYPE_DOCUMENT, [
      { name: "word/styles.xml", type: STYLES_W, body: "<w:styles/>" },
      { name: "word/theme/theme1.xml", type: THEME, body: "<a:theme/>" },
      { name: "word/footnotes.xml", type: FOOTNOTES, body: "<w:footnotes/>" },
    ])
    const pkg = await openPackage(bytes, "Report.docx", WORD_PACKAGE)
    expect(pkg.mainPartName).toBe("/word/document.xml")
    const names = (await readZip(await writePackage(pkg, WORD_PACKAGE))).map((e) => e.normalizedName)
    expect(names.slice(0, 3)).toEqual(["[Content_Types].xml", "_rels/.rels", "word/document.xml"])
    expect(names.indexOf("word/footnotes.xml")).toBeLessThan(names.indexOf("word/theme/theme1.xml"))
    expect(names.indexOf("word/theme/theme1.xml")).toBeLessThan(names.indexOf("word/styles.xml"))
  })

  it("opens and writes a workbook", async () => {
    const bytes = await minimal("xl/workbook.xml", CONTENT_TYPE_WORKBOOK)
    const pkg = await openPackage(bytes, "Book1.xlsx", SPREADSHEET_PACKAGE)
    const again = await openPackage(await writePackage(pkg, SPREADSHEET_PACKAGE), "Book1.xlsx", SPREADSHEET_PACKAGE)
    expect(again.mainPartName).toBe("/xl/workbook.xml")
  })

  it("names the format in its errors", async () => {
    const doc = await minimal("word/document.xml", CONTENT_TYPE_DOCUMENT)
    await expect(openPackage(doc, "Book1.xlsx", SPREADSHEET_PACKAGE)).rejects.toMatchObject({
      code: "content-type-mismatch",
    })
    await expect(openPackage(doc, "notes.zip", WORD_PACKAGE)).rejects.toThrow("as a Word document")
    await expect(openPackage(doc, "notes.zip", SPREADSHEET_PACKAGE)).rejects.toThrow("as an Excel workbook")
    await expect(openPackage(doc, "notes.zip", WORD_PACKAGE)).rejects.toBeInstanceOf(OpcError)
  })

  it("knows each format's extensions, main types and macros", () => {
    expect(packageExtensions(WORD_PACKAGE)).toEqual(["docx", "docm", "dotx", "dotm"])
    expect(packageExtensions(SPREADSHEET_PACKAGE)).toEqual(["xlsx", "xlsm", "xltx", "xltm"])
    expect(mainContentTypeFor("a.txt", WORD_PACKAGE)).toBe(CONTENT_TYPE_DOCUMENT)
    expect(mainContentTypeFor("a.xlsx", SPREADSHEET_PACKAGE)).toBe(CONTENT_TYPE_WORKBOOK)
    expect(allowsMacros("C:\\Files\\Report.DOCM", WORD_PACKAGE)).toBe(true)
    expect(allowsMacros("Book.xlsm", WORD_PACKAGE)).toBe(false)
    expect(allowsMacros("Book.xlsm", SPREADSHEET_PACKAGE)).toBe(true)
  })

  it("keeps the app's passthrough parts", () => {
    const part = (contentType: string, name = "/x.xml") => ({ name, contentType, data: new Uint8Array() })
    const fontTable = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    expect(isPassthroughPart(part(fontTable), WORD_PACKAGE)).toBe(true)
    expect(isPassthroughPart(part(fontTable), SPREADSHEET_PACKAGE)).toBe(false)
    expect(isPassthroughPart(part(THEME), SPREADSHEET_PACKAGE)).toBe(true)
    expect(isPassthroughPart(part("application/xml", "/customXml/item1.xml"), WORD_PACKAGE)).toBe(true)
  })

  it("lets an app describe its own format", async () => {
    const custom: PackageFormat = { ...SPREADSHEET_PACKAGE, partRank: (n) => (/styles/.test(n) ? 5 : undefined) }
    const bytes = await minimal("xl/workbook.xml", CONTENT_TYPE_WORKBOOK, [
      { name: "xl/worksheets/sheet1.xml", type: "application/xml", body: "<worksheet/>" },
      { name: "xl/styles.xml", type: "application/xml", body: "<styleSheet/>" },
    ])
    const pkg = await openPackage(bytes, "Book1.xlsx", custom)
    const names = (await readZip(await writePackage(pkg, custom))).map((e) => e.normalizedName)
    expect(names.indexOf("xl/styles.xml")).toBeLessThan(names.indexOf("xl/worksheets/sheet1.xml"))
  })
})

describe("zip and XML", () => {
  it("a stored zip (written without await) reads back", async () => {
    const entries = await readZip(storedZip([{ name: "a/ü.xml", data: text("<a>é</a>") }]))
    expect(entries.map((e) => e.normalizedName)).toEqual(["a/ü.xml"])
    expect(decodeXml(entries[0]!.data)).toBe("<a>é</a>")
  })

  it("escapes and parses", () => {
    const seen: string[] = []
    parseElements(`<r a="${escapeXml(`<"&">`)}"/>`, (e) => seen.push(e.attributes.a!))
    expect(seen).toEqual([`<"&">`])
  })
})
