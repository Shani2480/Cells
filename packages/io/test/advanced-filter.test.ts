// An Advanced Filter in place, as Excel 365 saves it (fixtures/excel/advanced-filter.xlsx: Sales!A1:B6 filtered by
// D1:D2, Qty > 5): sheetPr filterMode, the list as _xlnm._FilterDatabase and the criteria as _xlnm.Criteria. Excel
// keeps no "unique records only" flag.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { Workbook } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/advanced-filter.xlsx", import.meta.url)))

const text = async (bytes: Uint8Array, part: string) => {
  const zip = await readZip(bytes)
  return new TextDecoder().decode(zip.find((e) => e.normalizedName === part)!.data)
}

const LIST = { r0: 0, c0: 0, r1: 5, c1: 1 }
const CRITERIA = { sheet: 0, range: { r0: 0, c0: 3, r1: 1, c1: 3 } }

describe("Advanced Filter in place, in xlsx", () => {
  it("reads Excel's, so Clear and Reapply know it", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "advanced-filter.xlsx")
    const s = wb.sheets[0]
    expect(s.advancedFilter).toEqual({ list: LIST, criteria: CRITERIA, unique: false })
    expect(s.autoFilter).toBeNull()
    expect(s.rows.isHidden(1)).toBe(true)
    // The criteria name is the filter's, not one of the workbook's names.
    expect(wb.names.all().map((n) => n.name)).not.toContain("_xlnm.Criteria")
  })

  it("writes it as Excel does, and reads it back", async () => {
    const { workbook: wb, source } = await openXlsx(fixture(), "advanced-filter.xlsx")
    const bytes = await saveXlsx(wb, { source, fileName: "advanced-filter.xlsx" })
    const book = await text(bytes, "xl/workbook.xml")
    expect(book).toContain(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sales!$A$1:$B$6</definedName>`,
    )
    expect(book).toContain(`<definedName name="_xlnm.Criteria" localSheetId="0">Sales!$D$1:$D$2</definedName>`)
    expect(book.match(/_xlnm\.Criteria/g)).toHaveLength(1)
    expect(await text(bytes, "xl/worksheets/sheet1.xml")).toMatch(/<sheetPr[^>]*filterMode="1"/)
    const again = (await openXlsx(bytes, "again.xlsx")).workbook.sheets[0]
    expect(again.advancedFilter).toEqual({ list: LIST, criteria: CRITERIA, unique: false })
  })

  it("a new workbook's filter, with criteria on another sheet, and none once it's cleared", async () => {
    const wb = new Workbook(["Sales", "Rules"])
    wb.sheets[0].cells.set(0, 0, { value: "Qty" })
    wb.sheets[0].advancedFilter = {
      list: { r0: 0, c0: 0, r1: 4, c1: 0 },
      criteria: { sheet: 1, range: { r0: 0, c0: 0, r1: 1, c1: 0 } },
      unique: true,
    }
    const bytes = await saveXlsx(wb, { fileName: "new.xlsx" })
    expect(await text(bytes, "xl/workbook.xml")).toContain(
      `<definedName name="_xlnm.Criteria" localSheetId="0">Rules!$A$1:$A$2</definedName>`,
    )
    const read = (await openXlsx(bytes, "new.xlsx")).workbook.sheets[0]
    expect(read.advancedFilter?.criteria).toEqual({ sheet: 1, range: { r0: 0, c0: 0, r1: 1, c1: 0 } })

    const { workbook, source } = await openXlsx(fixture(), "advanced-filter.xlsx")
    workbook.sheets[0].advancedFilter = null
    const cleared = await saveXlsx(workbook, { source, fileName: "cleared.xlsx" })
    expect(await text(cleared, "xl/workbook.xml")).not.toMatch(/Criteria|_FilterDatabase/)
    expect(await text(cleared, "xl/worksheets/sheet1.xml")).not.toMatch(/filterMode/)
  })
})
