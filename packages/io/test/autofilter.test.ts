// AutoFilter in xlsx: a file Excel 365 saved with one filter per sheet (fixtures/excel/filters.xlsx), read and
// written back as Excel writes it.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/filters.xlsx", import.meta.url)))
const range = { r0: 0, c0: 0, r1: 6, c1: 2 }

const text = async (bytes: Uint8Array, part: string) => {
  const zip = await readZip(bytes)
  return new TextDecoder().decode(zip.find((e) => e.normalizedName === part)!.data)
}

describe("AutoFilter in xlsx", () => {
  it("reads every kind Excel wrote", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "filters.xlsx")
    const f = wb.sheets.map((s) => s.autoFilter)
    expect(f[0]).toEqual({ range, columns: { 0: { kind: "values", values: ["Ann", "Cy"] } } })
    expect(f[1]?.columns[1]).toEqual({
      kind: "custom",
      conditions: [
        { op: "ge", value: "3" },
        { op: "lt", value: "8" },
      ],
    })
    expect(f[2]?.columns[0]).toEqual({
      kind: "custom",
      and: false,
      conditions: [
        { op: "eq", value: "B*" },
        { op: "eq", value: "*e" },
      ],
    })
    expect(f[3]?.columns[1]).toEqual({ kind: "top10", top: true, percent: false, count: 2 })
    expect(f[4]?.columns[2]).toEqual({
      kind: "values",
      values: [],
      dates: [
        { year: 2026, month: 3, day: 1 },
        { year: 2026, month: 4 },
      ],
    })
    expect(f[5]?.columns[0]).toEqual({ kind: "color", cell: true, color: "#FF0000" })
    expect(f[6]).toEqual({ range, columns: {}, sort: { col: 1, descending: true } })
    // The rows Excel hid stay hidden.
    expect([1, 2, 3, 4, 5, 6].map((r) => wb.sheets[0].rows.isHidden(r))).toEqual([false, true, false, true, true, true])
  })

  it("writes them back as Excel does, with the hidden _FilterDatabase names", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "filters.xlsx")
    const saved = await saveXlsx(wb)
    const sheet = async (i: number) =>
      (await text(saved, `xl/worksheets/sheet${i}.xml`)).match(
        /<autoFilter[^>]*\/>|<autoFilter[\s\S]*?<\/autoFilter>/,
      )![0]
    expect(await sheet(1)).toBe(
      `<autoFilter ref="A1:C7"><filterColumn colId="0"><filters><filter val="Ann"/><filter val="Cy"/></filters></filterColumn></autoFilter>`,
    )
    expect(await sheet(2)).toContain(
      `<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="3"/><customFilter operator="lessThan" val="8"/></customFilters>`,
    )
    expect(await sheet(3)).toContain(`<customFilters><customFilter val="B*"/><customFilter val="*e"/></customFilters>`)
    expect(await sheet(4)).toContain(`<top10 val="2"/>`)
    expect(await sheet(5)).toContain(
      `<filters><dateGroupItem year="2026" month="3" day="1" dateTimeGrouping="day"/><dateGroupItem year="2026" month="4" dateTimeGrouping="month"/></filters>`,
    )
    expect(await sheet(6)).toMatch(/<colorFilter dxfId="\d+"\/>/)
    expect(await sheet(7)).toBe(
      `<autoFilter ref="A1:C7"><sortState ref="A2:C7"><sortCondition descending="1" ref="B2:B7"/></sortState></autoFilter>`,
    )
    const book = await text(saved, "xl/workbook.xml")
    expect(book.match(/_xlnm\._FilterDatabase/g)).toHaveLength(7)
    expect(book).toContain(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$C$7</definedName>`,
    )
    // And it reads back the same.
    const again = await openXlsx(saved, "filters.xlsx")
    expect(again.workbook.sheets.map((s) => s.autoFilter)).toEqual(wb.sheets.map((s) => s.autoFilter))
  })

  it("a new workbook's filter on a sheet with a quoted name", async () => {
    const wb = new Workbook(["Q1 sales"])
    wb.sheets[0].cells.set(0, 0, { value: "x" })
    wb.sheets[0].autoFilter = {
      range: { r0: 0, c0: 0, r1: 3, c1: 0 },
      columns: { 0: { kind: "color", cell: false, color: "#0000FF" } },
    }
    const saved = await saveXlsx(wb)
    expect(await text(saved, "xl/workbook.xml")).toContain(`hidden="1">'Q1 sales'!$A$1:$A$4</definedName>`)
    const again = await openXlsx(saved, "x.xlsx")
    expect(again.workbook.sheets[0].autoFilter?.columns[0]).toEqual({ kind: "color", cell: false, color: "#0000FF" })
  })
})
