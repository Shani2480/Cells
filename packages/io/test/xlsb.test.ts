import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { openXlsx } from "../src/xlsx/read"
import { openXlsb, XlsbError } from "../src/xlsb/read"
import { snapshot } from "./workbook-snapshot"

const file = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/xlsb/${name}`, import.meta.url)))

describe(".xlsb (Excel Binary Workbook)", () => {
  it("opens as Excel's .xlsx of the same workbook does: cells, formulas, formats, sizes, merges, notes, links, names, panes", async () => {
    const xlsb = await openXlsb(file("demo.xlsb"), "demo.xlsb")
    const xlsx = await openXlsx(file("demo.xlsx"), "demo.xlsx")
    expect(xlsb.warnings).toEqual([])
    expect(snapshot(xlsb.workbook, xlsb.views)).toEqual(snapshot(xlsx.workbook, xlsx.views))
    const sales = xlsb.workbook.sheets[0]
    expect(sales.cells.get(3, 5)?.input).toBe("=E4*VLOOKUP(B4,Rates!$A$2:$B$3,2,FALSE)")
    expect(sales.cells.get(5, 5)?.input).toBe('=TEXTJOIN(", ",TRUE,B4:B6)')
    expect(sales.cells.get(8, 8)?.input).toBe("=SUM(Units)")
    expect(xlsb.activeSheet).toBe(0)
  })

  it("says why a file isn't one: password-protected (a compound file) or not a package", async () => {
    const cfb = new Uint8Array(512)
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    await expect(openXlsb(cfb, "secret.xlsb")).rejects.toMatchObject({ kind: "encrypted" })
    await expect(openXlsb(new Uint8Array([1, 2, 3]), "junk.xlsb")).rejects.toBeInstanceOf(XlsbError)
  })
})

describe(".xlsb: tables, shared and array formulas, newer functions", () => {
  it("opens as Excel's .xlsx of the same workbook does", async () => {
    const xlsb = await openXlsb(file("more.xlsb"), "more.xlsb")
    const xlsx = await openXlsx(file("more.xlsx"), "more.xlsx")
    expect(xlsb.warnings).toEqual(xlsx.warnings)
    expect(snapshot(xlsb.workbook, xlsb.views, 16, 16)).toEqual(snapshot(xlsx.workbook, xlsx.views, 16, 16))
    const formulas = (wb: typeof xlsb.workbook) => {
      const out: string[] = []
      wb.sheets.forEach((s) => s.cells.forEachFormula((r, c, input) => out.push(`${s.name}!${r},${c} ${input}`)))
      return out
    }
    expect(formulas(xlsb.workbook)).toEqual(formulas(xlsx.workbook))
    const tables = (wb: typeof xlsb.workbook) =>
      wb.sheets.flatMap((s) =>
        s.tables.map((t) => ({
          name: t.name,
          range: t.range,
          columns: t.columns.map((c) => c.name),
          totals: t.totals,
          header: t.header,
          style: t.style,
        })),
      )
    expect(tables(xlsb.workbook)).toEqual(tables(xlsx.workbook))
    expect(tables(xlsb.workbook)).toHaveLength(1)
    const sales = xlsb.workbook.sheets[0]
    expect(sales.cells.get(0, 10)?.input).toBe("=SEQUENCE(3)")
    expect(sales.cells.get(2, 10)?.value).toBe(3)
    expect(xlsb.workbook.sheets[2].hidden).toBe(true)
    expect(sales.pageSetup).toEqual(xlsx.workbook.sheets[0].pageSetup)
  })
})
