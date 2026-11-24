// A workbook's own table style in xlsx: the file Excel 365 made (fixtures/excel/table-custom-style.xlsx: a made-up
// parts list as a table in "Table Style 1", the workbook's default table style, with a total row).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { tableCellLook, tableStyleDef } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}`, import.meta.url)))
const stylesXml = async (bytes: Uint8Array) => {
  const zip = await readZip(bytes)
  const part = zip.find((e) => e.normalizedName.endsWith("styles.xml"))!
  return new TextDecoder().decode(part.data)
}

describe("custom table styles in xlsx", () => {
  it("reads a table style Excel made, draws the table with it and writes it back (table-custom-style.xlsx)", async () => {
    const read = await openXlsx(fixture("table-custom-style.xlsx"), "table-custom-style.xlsx")
    const wb = read.workbook
    expect(wb.tableStyles.defaultTable).toBe("Table Style 1")
    expect(wb.tableStyles.defaultPivot).toBeNull()
    const [style] = wb.tableStyles.styles
    expect(style).toMatchObject({ name: "Table Style 1", table: true, pivot: false })
    expect(style.elements).toEqual({
      wholeTable: { border: { bottom: { style: "thin", color: "#404040" } } },
      headerRow: { fill: "#1F3F7F" },
      totalRow: { bold: false, italic: true, border: { top: { style: "double" } } },
      firstRowStripe: { fill: "#DAE6F2", stripe: 2 },
    })
    const table = wb.sheets[0].tables[0]
    expect(table.style).toBe("Table Style 1")
    const def = tableStyleDef(table.style, wb.theme, wb.tableStyles.styles)
    const look = (row: number, col: number) => tableCellLook(table, row, col, def)
    expect(look(0, 0)).toMatchObject({ fill: "#1F3F7F" })
    // A first stripe of two rows and a second of one: two shaded, one not, two shaded (what Excel 365 draws).
    expect([1, 2, 3, 4, 5].map((r) => look(r, 1)?.fill ?? null)).toEqual([
      "#DAE6F2",
      "#DAE6F2",
      null,
      "#DAE6F2",
      "#DAE6F2",
    ])
    expect(look(table.range.r1, 0)).toMatchObject({ italic: true, borderTop: { style: "double" } })
    // Not a PivotTable style: the PivotTable gallery doesn't offer it.
    expect(tableStyleDef("Table Style 1", wb.theme, [{ ...style, table: false }])).toBeNull()

    expect(await stylesXml(await saveXlsx(wb, { source: read.source }))).toContain(
      'xr9:uid="{392434EB-A09E-4518-8A7E-278340CDB2C5}"',
    )
    wb.tableStyles = { ...wb.tableStyles, styles: [{ ...style, name: "Parts" }], defaultTable: "Parts" }
    const out = await saveXlsx(wb, { source: read.source })
    const xml = await stylesXml(out)
    expect(xml).toMatch(/defaultTableStyle="Parts" defaultPivotStyle="PivotStyleLight16"/)
    expect(xml).toMatch(/<tableStyle name="Parts" pivot="0" count="4">/)
    expect((await openXlsx(out, "again.xlsx")).workbook.tableStyles).toEqual(wb.tableStyles)
  })
})
