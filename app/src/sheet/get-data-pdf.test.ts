import { readFileSync } from "node:fs"
import { join } from "node:path"
import { EN_US, readQueryPlan, setRegional, type Regional } from "@cells/core"
import { afterEach, describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { hasHeadings, setFileReader } from "./get-data"
import { dateFrom, numberFrom } from "./power-query"

const fixture = (name: string) =>
  new Uint8Array(readFileSync(join(__dirname, "../../../packages/io/test/fixtures/pdf", name)))

const GB: Regional = { decimal: ".", group: ",", dateSeparator: "/", dateOrder: "dmy", shortDate: "dd/mm/yyyy" }
const NL: Regional = { decimal: ",", group: ".", dateSeparator: "-", dateOrder: "dmy", shortDate: "d-m-yyyy" }
const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000

afterEach(() => setFileReader(async () => Promise.reject(new Error("no file"))))

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
