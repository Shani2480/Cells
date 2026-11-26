import { readFileSync } from "node:fs"
import { join } from "node:path"
import { EN_US, readQueryPlan, setRegional, type Regional } from "@cells/core"
import { pdfFromImages, type PdfRun } from "@cells/io"
import { afterEach, describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"
import type { Recognizer } from "@/pdf/pdf"
import { hasHeadings, setFileReader, setRecognitionSwitch, setRecognizer } from "./get-data"
import { dateFrom, numberFrom } from "./power-query"

const fixture = (name: string) =>
  new Uint8Array(readFileSync(join(__dirname, "../../../packages/io/test/fixtures/pdf", name)))

const GB: Regional = { decimal: ".", group: ",", dateSeparator: "/", dateOrder: "dmy", shortDate: "dd/mm/yyyy" }
const NL: Regional = { decimal: ",", group: ".", dateSeparator: "-", dateOrder: "dmy", shortDate: "d-m-yyyy" }
const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000

afterEach(() => {
  setFileReader(async () => Promise.reject(new Error("no file")))
  setRecognizer(null)
})

describe("text as numbers and dates, as Power Query reads them", () => {
  it("reads numbers with the region's separators, signs in brackets or after, and percentages", () => {
    expect(numberFrom("1,200.00", EN_US)).toBe(1200)
    expect(numberFrom("(5.00)", EN_US)).toBe(-5)
    expect(numberFrom("5.00-", EN_US)).toBe(-5)
    expect(numberFrom("12%", EN_US)).toBeCloseTo(0.12)
    expect(numberFrom("12,50", NL)).toBe(12.5)
    expect(numberFrom("1.200,50", NL)).toBe(1200.5)
    // A PDF's currency signs and spaced groups too.
    expect(numberFrom("€ 450.00", EN_US)).toBe(450)
    expect(numberFrom("-$1,200", EN_US)).toBe(-1200)
    expect(numberFrom("120 500", EN_US)).toBe(120500)
    expect(numberFrom("A-100", EN_US)).toBeNull()
    expect(numberFrom("2026-01-03", EN_US)).toBeNull()
  })

  it("reads dates in the region's order, ISO dates and month names always", () => {
    expect(dateFrom("01/02/2026", EN_US)).toBe(serial(2026, 1, 2))
    expect(dateFrom("01/02/2026", GB)).toBe(serial(2026, 2, 1))
    expect(dateFrom("13/02/2026", EN_US)).toBeNull()
    expect(dateFrom("13.02.2026", GB)).toBe(serial(2026, 2, 13))
    expect(dateFrom("2026-02-13", EN_US)).toBe(serial(2026, 2, 13))
    expect(dateFrom("13 Feb 2026", EN_US)).toBe(serial(2026, 2, 13))
    expect(dateFrom("Feb 13, 2026", GB)).toBe(serial(2026, 2, 13))
    expect(dateFrom("31/02/2026", GB)).toBeNull()
  })

  it("sees a table's headings over numbers", () => {
    const t = (rows: (string | null)[][]) => ({ columns: [], rows, types: [] })
    expect(
      hasHeadings(
        t([
          ["Item", "Qty"],
          ["Pens", "4"],
        ]),
      ),
    ).toBe(true)
    expect(
      hasHeadings(
        t([
          ["Pens", "4"],
          ["Ink", "2"],
        ]),
      ),
    ).toBe(false)
    expect(
      hasHeadings(
        t([
          ["Item", "Notes"],
          ["Pens", "blue"],
        ]),
      ),
    ).toBe(false)
  })
})

describe("Get Data › From PDF", () => {
  it("lists the tables and pages, and loads a table typed in the person's region", async () => {
    setRegional(GB)
    try {
      const c = new SheetController()
      await c.getData.fromPdf("/home/alex/Documents/statement.pdf", fixture("statement.pdf"))
      const nav = c.getData.navigator
      expect(nav?.title).toBe("statement.pdf")
      expect(nav?.items.map((i) => i.name)).toEqual(["Table001 (Page 1-3)", "Page001", "Page002", "Page003"])
      await c.getData.loadItems([0])
      const query = c.wb.queries[0]
      expect(query.name).toBe("Table001")
      const plan = readQueryPlan(query.formula)
      expect(plan?.source).toEqual({ kind: "pdf", path: "/home/alex/Documents/statement.pdf" })
      expect(plan?.steps.map((s) => s.name)).toEqual(["Table001", "Promoted Headers", "Changed Type"])
      expect(plan?.steps[2].step).toEqual({
        op: "types",
        columns: [
          ["Date", "date"],
          ["Description", "text"],
          ["Amount", "number"],
          ["Balance", "number"],
        ],
      })
      const sheet = c.wb.sheets[c.sheetIndex]
      const table = sheet.tables[0]
      expect(table.query).toBe("Table001")
      expect(table.range.r1 - table.range.r0).toBe(90)
      const cell = (r: number, col: number) => sheet.cells.value(table.range.r0 + r, table.range.c0 + col)
      expect(cell(0, 0)).toBe("Date")
      expect(cell(1, 0)).toBe(serial(2026, 1, 1))
      expect(cell(3, 2)).toBe(2500)
      expect(cell(90, 3)).toBe(18376)
    } finally {
      setRegional(EN_US)
    }
  })

  it("reads another region's dates with that locale, as Excel's Using Locale", async () => {
    const c = new SheetController()
    await c.getData.fromPdf("/home/alex/statement.pdf", fixture("statement.pdf"))
    await c.getData.loadItems([0])
    const plan = readQueryPlan(c.wb.queries[0].formula)
    expect(plan?.steps.slice(2)).toEqual([
      {
        name: "Changed Type",
        step: {
          op: "types",
          columns: [
            ["Description", "text"],
            ["Amount", "number"],
            ["Balance", "number"],
          ],
        },
      },
      { name: "Changed Type with Locale", step: { op: "types", columns: [["Date", "date"]], culture: "en-GB" } },
    ])
    expect(c.wb.queries[0].formula).toContain('{{"Date", type date}}, "en-GB")')
    const sheet = c.wb.sheets[c.sheetIndex]
    expect(sheet.cells.value(sheet.tables[0].range.r0 + 13, sheet.tables[0].range.c0)).toBe(serial(2026, 1, 13))
  })

  it("loads several items at once, and refreshes from the file", async () => {
    const c = new SheetController()
    let reads = 0
    setFileReader(async () => {
      reads++
      return fixture("rules.pdf")
    })
    await c.getData.fromPdf("/home/alex/rules.pdf", fixture("rules.pdf"))
    await c.getData.loadItems([0, 1])
    expect(c.wb.queries.map((q) => q.name)).toEqual(["Table001", "Table002"])
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "Table001", "Table002"])
    expect(reads).toBe(0)
    expect(await c.getData.refresh("Table002")).toBe(true)
    expect(reads).toBe(1)
  })

  it("says so when the file can't be read again", async () => {
    const c = new SheetController()
    const q = `let\r\n    Source = Pdf.Tables(File.Contents("/home/alex/gone.pdf"), [Implementation="1.3"]),\r\n    Table001 = Source{[Id="Table001"]}[Data]\r\nin\r\n    Table001`
    c.wb.transact([{ kind: "queries", queries: [{ name: "Gone", description: "", formula: q, table: null }] }], {
      label: "Q",
    })
    await expect(c.getData.loadConnection("Gone")).rejects.toThrow("Cells couldn't read gone.pdf")
  })
})

describe("Get Data › From PDF, scanned", () => {
  // The demo invoice scanned at 300 dpi, as a PDF of that one picture.
  const jpeg = fixture("scan-invoice.jpg")
  const scan = pdfFromImages([
    { width: (2507 * 72) / 300, height: (3528 * 72) / 300, jpeg, pixelWidth: 2507, pixelHeight: 3528 },
  ])
  const run = (text: string, x0: number, y: number, unsure = false): PdfRun => ({
    text,
    x0,
    x1: x0 + text.length * 5,
    y0: y,
    y1: y + 10,
    size: 10,
    words: [{ text, x0, x1: x0 + text.length * 5, ...(unsure ? { unsure } : {}) }],
  })
  // Stands in for Tesseract (ocr.test.ts reads the real scans): a heading and two rows, one figure unsure.
  const fake: Recognizer = async (page) => ({
    ...page,
    runs: [
      run("Item", 60, 100),
      run("Qty", 300, 100),
      run("A-210", 60, 120),
      run("500", 300, 120),
      run("C-300", 60, 140),
      run("25", 300, 140, true),
    ],
  })

  it("asks to turn text recognition on, then reads the scan and marks what it wasn't sure of", async () => {
    const c = new SheetController()
    const asked: string[] = []
    c.shell = { ...c.shell, confirm: (m, ok) => (asked.push(m), ok()) }
    const busy: (string | null)[] = []
    c.subscribe(() => busy.push(c.getData.busy))
    setRecognitionSwitch(() => setRecognizer(fake))
    await c.getData.fromPdf("/home/alex/Scans/invoice.pdf", scan)
    expect(asked[0]).toMatch(/^This PDF is scanned/)
    await vi.waitFor(() => expect(c.getData.navigator).not.toBeNull())
    const nav = c.getData.navigator
    expect(nav?.unread).toEqual([])
    expect(nav?.items[0].preview?.columns).toEqual(["Item", "Qty"])
    expect(nav?.items[0].preview?.rows).toEqual([
      ["A-210", 500],
      ["C-300", 25],
    ])
    // The heading promoted, the unsure figure is the second row's.
    expect(nav?.items[0].unsure).toEqual([[1, 1]])
    expect(busy).toContain("Recognizing text…")
    expect(c.getData.busy).toBeNull()
  })

  it("says why when text recognition isn't available", async () => {
    const c = new SheetController()
    setRecognitionSwitch(null)
    await expect(c.getData.fromPdf("/home/alex/Scans/invoice.pdf", scan)).rejects.toThrow(
      "text recognition isn't available here",
    )
  })
})
