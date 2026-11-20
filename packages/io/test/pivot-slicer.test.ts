// Slicers on PivotTables that Excel 365 made (fixtures/excel/pivot-slicer.xlsx): Product and Channel filter PTOne
// (Plums unpicked; neither field is on PTOne's axes), Region filters PTTwo and PTTwoB (East unpicked; Region is
// PTTwo's row field and off PTTwoB's), with the Dark 2 style and two columns.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  isChart,
  pivotLayout,
  pivotSlicerItems,
  pivotSlicerTarget,
  withShownItems,
  type SheetSlicer,
  type Workbook,
} from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/pivot-slicer.xlsx", import.meta.url)))
async function parts(bytes: Uint8Array) {
  const zip = await readZip(bytes)
  return new Map(zip.map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
}
const slicers = (wb: Workbook) =>
  wb.sheets.flatMap((s) => s.objects.filter((o): o is SheetSlicer => !isChart(o) && o.kind === "slicer"))
const total = (wb: Workbook, name: string) => {
  const table = wb.sheets.flatMap((s) => s.pivots).find((t) => t.name === name)!
  const cache = wb.pivotCaches.find((c) => c.id === table.cacheId)!
  const cells = pivotLayout(table, cache).cells
  const label = cells.find((c) => c.value === "Grand Total")!
  return cells.find((c) => c.row === label.row && typeof c.value === "number")!.value
}
const picked = (wb: Workbook, o: SheetSlicer) =>
  pivotSlicerItems(pivotSlicerTarget(wb.sheets, wb.pivotCaches, o)!, o)
    .filter((i) => i.selected)
    .map((i) => i.text)

describe("PivotTable slicers", () => {
  it("reads Excel's slicers: their PivotTables, field and picks, and the PivotTables filtered by them", async () => {
    const { workbook: wb, warnings } = await openXlsx(fixture(), "pivot-slicer.xlsx")
    expect(warnings).toEqual([])
    const [product, channel, region] = slicers(wb)
    expect(product).toMatchObject({ name: "Product", column: "Product", caption: "Product" })
    expect(product.pivot).toMatchObject({ field: "Product", tables: [{ sheet: "One", name: "PTOne" }] })
    expect(channel.caption).toBe("Sales channel")
    expect(region).toMatchObject({ style: "SlicerStyleDark2", columns: 2 })
    expect(region.pivot?.tables).toEqual([
      { sheet: "Two", name: "PTTwo" },
      { sheet: "Two", name: "PTTwoB" },
    ])
    expect(picked(wb, product)).toEqual(["Apples", "Pears"])
    expect(picked(wb, region)).toEqual(["North", "South", "West"])
    // Excel's totals: Plums left out of PTOne (130 of 212 units), East out of PTTwoB (137) though Region is not on it.
    expect(total(wb, "PTOne")).toBe(130)
    expect(total(wb, "PTTwo")).toBe(4707.5)
    expect(total(wb, "PTTwoB")).toBe(137)
  })

  it("writes them back as Excel does, and a changed pick is the PivotTables' and the slicer cache's", async () => {
    const read = await openXlsx(fixture(), "pivot-slicer.xlsx")
    const wb = read.workbook
    // Pears only on Product, as a slicer click does it.
    const one = wb.sheets[1]
    const cache = wb.pivotCaches[0]
    one.pivots = one.pivots.map((t) => withShownItems(t, cache, 1, new Set([2])))
    const bytes = await saveXlsx(wb, { source: read.source, fileName: "pivot-slicer.xlsx" })
    const out = await parts(bytes)
    const wbXml = out.get("xl/workbook.xml")!
    expect(wbXml).toContain('<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}"')
    expect(wbXml).toContain('<definedName name="Slicer_Region">#N/A</definedName>')
    const caches = [...out].filter(([k]) => k.startsWith("xl/slicerCaches/")).map(([, v]) => v)
    expect(caches).toHaveLength(3)
    const product = caches.find((c) => c.includes('sourceName="Product"'))!
    expect(product).toContain('<pivotTables><pivotTable tabId="2" name="PTOne"/></pivotTables>')
    expect(product).toMatch(
      /<tabular pivotCacheId="645199731"><items count="3"><i x="0"\/><i x="2" s="1"\/><i x="1"\/>/,
    )
    const region = caches.find((c) => c.includes('sourceName="Region"'))!
    expect(region).toContain('<pivotTable tabId="3" name="PTTwo"/><pivotTable tabId="3" name="PTTwoB"/>')
    const sheets = [...out].filter(([k]) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)).map(([, v]) => v)
    expect(sheets.filter((x) => x.includes("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"))).toHaveLength(2)

    const { workbook: again, warnings } = await openXlsx(bytes, "pivot-slicer.xlsx")
    expect(warnings).toEqual([])
    const [p, , r] = slicers(again)
    expect(picked(again, p)).toEqual(["Pears"])
    expect(picked(again, r)).toEqual(["North", "South", "West"])
    expect(r).toMatchObject({ style: "SlicerStyleDark2", columns: 2 })
    expect(total(again, "PTOne")).toBe(67)
    expect(total(again, "PTTwoB")).toBe(137)
  })
})
