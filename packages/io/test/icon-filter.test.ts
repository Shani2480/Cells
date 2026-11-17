// Filter and sort by conditional formatting icon, checked against a file Excel 365 made (fixtures/excel/
// iconfilter.xlsx): Sheet1 filters the green up arrow, Sheet2 No Cell Icon, Sheet3 was sorted red down arrow on top.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { ConditionalEval, filterVisibility, sortOps, applyOp, type CellLook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/iconfilter.xlsx", import.meta.url)))
const TODAY = 46000

describe("filter and sort by cell icon", () => {
  it("reads Excel's icon filters and shows the rows Excel showed", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "iconfilter.xlsx")
    const [s1, s2] = wb.sheets
    expect(s1.autoFilter?.columns[0]).toEqual({ kind: "icon", icon: { set: "3Arrows", index: 2 } })
    expect(s2.autoFilter?.columns[0]).toEqual({ kind: "icon", icon: null })
    for (const [i, s] of [s1, s2].entries()) {
      const ev = new ConditionalEval(wb.sheets, i, TODAY)
      const look: CellLook = (r, c) => ev.at(r, c)
      const shown = filterVisibility(s, wb.styles, s.autoFilter!, TODAY, look)
      // What Excel hid in the file.
      expect(shown).toEqual(shown.map((_, k) => !s.rows.isHidden(k + 1)))
    }
    expect(filterVisibility(s1, wb.styles, s1.autoFilter!, TODAY).some(Boolean)).toBe(false)
  })

  it("sorts an icon on top as Excel did", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "iconfilter.xlsx")
    const s = wb.sheets[2]
    const sorted = [1, 2, 3, 4, 5, 6, 7, 8, 9].map((r) => s.cells.get(r, 0)?.value)
    // Put the values back in the order Excel had them before its sort.
    ;[1, 9, 5, 2, 8, 4, 7, 3, 6].forEach((v, i) => s.cells.set(i + 1, 0, { value: v }))
    const ev = new ConditionalEval(wb.sheets, 2, TODAY)
    const range = { r0: 0, c0: 0, r1: 9, c1: 0 }
    const ops = sortOps(
      2,
      s,
      wb.styles,
      range,
      { header: true, keys: [{ index: 0, on: "icon", icon: { set: "3Arrows", index: 0 } }] },
      (r, c) => ev.at(r, c),
    )
    for (const op of ops) applyOp(wb.sheets, op, wb.styles)
    expect([1, 2, 3, 4, 5, 6, 7, 8, 9].map((r) => s.cells.get(r, 0)?.value)).toEqual(sorted)
  })

  it("writes them back as Excel does", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "iconfilter.xlsx")
    const zip = await readZip(await saveXlsx(wb))
    const part = (i: number) =>
      new TextDecoder().decode(zip.find((e) => e.normalizedName === `xl/worksheets/sheet${i}.xml`)!.data)
    expect(part(1)).toContain(`<filterColumn colId="0"><iconFilter iconSet="3Arrows" iconId="2"/></filterColumn>`)
    expect(part(2)).toContain(`<filterColumn colId="0"><iconFilter iconSet="3Arrows"/></filterColumn>`)
  })
})
