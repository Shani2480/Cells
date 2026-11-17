// Table slicers in xlsx: a file Excel 365 made (fixtures/excel/slicer.xlsx): a table of regions, products and sales
// with a Region slicer (East and West unpicked) and a Product slicer (Light 2, two columns, its own caption).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { isChart, slicerItems, Workbook, type SheetSlicer } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/slicer.xlsx", import.meta.url)))
const slicersOf = (wb: Workbook) =>
  wb.sheets[0].objects.filter((o): o is SheetSlicer => !isChart(o) && o.kind === "slicer")
const part = async (bytes: Uint8Array, name: string) => {
  const zip = await readZip(bytes)
  const entry = zip.find((e) => e.normalizedName === name)
  return entry ? new TextDecoder().decode(entry.data) : null
}

describe("table slicers in xlsx", () => {
  it("reads each slicer, its table column and settings, and its picks from the table's filter", async () => {
    const { workbook: wb, warnings } = await openXlsx(fixture(), "slicer.xlsx")
    expect(warnings).toEqual([])
    const [region, product] = slicersOf(wb)
    expect(region).toMatchObject({
      name: "Region",
      caption: "Region",
      column: "Region",
      table: wb.sheets[0].tables[0].id,
      style: "SlicerStyleLight1",
      columns: 1,
      buttonHeight: 27,
      header: true,
      placement: "fixed",
      cache: "Slicer_Region",
    })
    expect(product).toMatchObject({ name: "Product 1", caption: "Product", style: "SlicerStyleLight2", columns: 2 })
    expect(wb.names.all().map((n) => n.name)).toEqual([])
    const items = slicerItems(wb.sheets, wb.styles, region, 46000)
    expect(items.map((i) => [i.text, i.selected, i.data])).toEqual([
      ["East", false, true],
      ["North", true, true],
      ["South", true, true],
      ["West", false, true],
    ])
    // Products: only the rows left by the Region filter have data.
    expect(slicerItems(wb.sheets, wb.styles, product, 46000).map((i) => [i.text, i.data])).toEqual([
      ["Apples", true],
      ["Pears", true],
      ["Plums", false],
    ])
  })

  it("writes the slicers, their caches and names anew, as Excel reads them", async () => {
    const read = await openXlsx(fixture(), "slicer.xlsx")
    const saved = await saveXlsx(read.workbook, { source: read.source })
    const wbXml = (await part(saved, "xl/workbook.xml"))!
    expect(wbXml.match(/<x14:slicerCache r:id=/g)).toHaveLength(2)
    expect(wbXml).toContain(`<definedName name="Slicer_Region">#N/A</definedName>`)
    expect(wbXml.match(/Slicer_Region/g)).toHaveLength(1)
    const sheet = (await part(saved, "xl/worksheets/sheet1.xml"))!
    expect(sheet.match(/<x14:slicerList/g)).toHaveLength(1)
    const slicers = (await part(saved, "xl/slicers/slicer1.xml"))!
    expect(slicers).toContain(`name="Product 1"`)
    expect(slicers).toContain(`columnCount="2"`)
    const drawing = (await part(saved, "xl/drawings/drawing1.xml"))!
    expect(drawing.match(/<sle:slicer /g)).toHaveLength(2)
    const again = await openXlsx(saved, "slicer.xlsx")
    expect(again.warnings).toEqual([])
    expect(slicersOf(again.workbook).map((s) => [s.name, s.column, s.style, s.columns])).toEqual([
      ["Region", "Region", "SlicerStyleLight1", 1],
      ["Product 1", "Product", "SlicerStyleLight2", 2],
    ])
    // Saved twice: still one cache each, no leftovers.
    const twice = await saveXlsx(again.workbook, { source: again.source })
    const zip = await readZip(twice)
    expect(
      zip
        .filter((e) => /slicer/i.test(e.normalizedName))
        .map((e) => e.normalizedName)
        .sort(),
    ).toEqual(["xl/slicerCaches/slicerCache1.xml", "xl/slicerCaches/slicerCache2.xml", "xl/slicers/slicer1.xml"])
  })

  it("leaves out a slicer whose table is gone", async () => {
    const read = await openXlsx(fixture(), "slicer.xlsx")
    read.workbook.sheets[0].tables = []
    const saved = await saveXlsx(read.workbook, { source: read.source })
    const zip = await readZip(saved)
    expect(zip.some((e) => /slicer/i.test(e.normalizedName))).toBe(false)
    expect(await part(saved, "xl/workbook.xml")).not.toContain("Slicer_")
  })
})
