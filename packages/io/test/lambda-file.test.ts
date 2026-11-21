// lambda.xlsx: LAMBDA in cells and in defined names, as Excel 365 saved it (made in the Office VM).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/lambda.xlsx", import.meta.url)))

const formulas = (xml: string) => [...xml.matchAll(/<f[^>]*>([^<]*)<\/f>/g)].map((m) => m[1])

describe("LAMBDA in xlsx", () => {
  it("opens Excel's file: named LAMBDAs are callable and every formula gives Excel's value", async () => {
    const { workbook: wb } = await openXlsx(bytes, "lambda.xlsx")
    wb.calculate(true)
    expect(wb.names.all().map((n) => [n.name, n.ref])).toEqual([
      ["AddUp", "LAMBDA(a,b,a+b)"],
      ["MyDouble", "LAMBDA(x,x*2)"],
    ])
    const v = (r: number, c = 0) => wb.sheets[0].cells.value(r, c)
    expect([v(0), v(1), v(2), v(3), v(3, 1), v(5), v(6), v(7), v(8), v(8, 1)]).toEqual([
      42, 6, 8, 10, 20, 1, 5, 3, 2, 3,
    ])
  })

  it("saves the formulas and names in Excel's own text", async () => {
    const { workbook: wb } = await openXlsx(bytes, "lambda.xlsx")
    const text = async (data: Uint8Array) =>
      new Map((await readZip(data)).map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
    const before = await text(bytes)
    const after = await text(await saveXlsx(wb))
    expect(formulas(after.get("xl/worksheets/sheet1.xml")!)).toEqual(formulas(before.get("xl/worksheets/sheet1.xml")!))
    const names = (xml: string) => [...xml.matchAll(/<definedName[^>]*>([^<]*)</g)].map((m) => m[1])
    expect(names(after.get("xl/workbook.xml")!)).toEqual(names(before.get("xl/workbook.xml")!))
  })
})
