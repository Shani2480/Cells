// PivotTables in xlsx: the files Excel 365 made (fixtures/excel/pivot-basic.xlsx and pivot-more.xlsx; the same
// made-up sales data with PivotTables in compact, outline and tabular form, crossed, with report filters, every
// summary function, custom subtotals, blank rows, collapsed items, captions and sorting).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  applyOp,
  cellName,
  newPivotField,
  pivotLayout,
  Workbook,
  type PivotCache,
  type PivotTable,
  type Sheet,
} from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"
import { pivotCacheKey, pivotTableKey } from "../src/xlsx/pivot"

const FILES = ["pivot-basic.xlsx", "pivot-more.xlsx"]
const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}`, import.meta.url)))

async function parts(bytes: Uint8Array) {
  const zip = await readZip(bytes)
  return new Map(zip.map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
}

const cacheOf = (wb: Workbook, t: PivotTable) => wb.pivotCaches.find((c) => c.id === t.cacheId)!

/** Each PivotTable laid out by the engine shows what Excel left in the sheet's cells. */
function expectCellsAsExcel(wb: Workbook) {
  let count = 0
  for (const sheet of wb.sheets)
    for (const t of sheet.pivots) {
      const layout = pivotLayout(t, cacheOf(wb, t))
      expect(layout.range, `${sheet.name} ${t.name}`).toEqual(t.location)
      for (const c of layout.cells) {
        const got = sheet.cells.get(c.row, c.col)?.value
        const want = c.value === undefined ? undefined : c.value
        const label = `${sheet.name}!${cellName({ row: c.row, col: c.col })}`
        if (typeof want === "number") expect(got, label).toBeCloseTo(want, 9)
        else expect(got ?? undefined, label).toEqual(want)
      }
      count++
    }
  return count
}

describe("PivotTables in xlsx", () => {
  it.each(FILES)("reads every PivotTable of %s, and the engine shows Excel's cells", async (name) => {
    const { workbook: wb } = await openXlsx(fixture(name), name)
    expect(wb.pivotCaches).toHaveLength(1)
    expect(wb.pivotCaches[0].records).toHaveLength(24)
    expect(wb.pivotCaches[0].source).toEqual({ kind: "range", sheet: "Data", ref: "A1:F25" })
    expect(expectCellsAsExcel(wb)).toBe(6)
  })

  it("reads the settings Excel wrote", async () => {
    const { workbook: wb } = await openXlsx(fixture("pivot-more.xlsx"), "pivot-more.xlsx")
    const all = wb.sheets.flatMap((s) => s.pivots)
    const byName = (n: string) => all.find((t) => t.name === n)!
    const collapsed = byName("PTCollapsed")
    expect(collapsed.captions.rowHeader).toBe("Where")
    expect(collapsed.fields[0].name).toBe("Area")
    expect(collapsed.fields[0].items.find((i) => i.collapsed)).toEqual({ x: 0, collapsed: true })
    const subs = byName("PTSubtotals")
    expect(subs.fields[5].subtotals).toEqual(["sum", "count"])
    expect(subs.fields[1].insertBlankRow).toBe(true)
    expect(subs.captions).toMatchObject({ error: "err", missing: "-" })
    expect(subs.showError).toBe(true)
    expect(byName("PTNumbers").fields[3].sort).toBe("descending")
    expect(byName("PTNumbers").captions.grandTotal).toBe("All of it")
    const pages = byName("PTTwoPages")
    expect(pages.pages).toEqual([
      { field: 5, item: null },
      { field: 0, item: null },
    ])
    expect(pages.fields[0].multipleItems).toBe(true)
    const { workbook: basic } = await openXlsx(fixture("pivot-basic.xlsx"), "pivot-basic.xlsx")
    const tabular = basic.sheets.flatMap((s) => s.pivots).find((t) => t.name === "PTTabular")!
    expect(tabular.fields[0]).toMatchObject({ outline: false, compact: false, repeatLabels: true, subtotals: [] })
    expect(tabular.colGrandTotals).toBe(false)
  })

  it.each(FILES)("writes %s's PivotTable parts back byte for byte when nothing changed", async (name) => {
    const read = await openXlsx(fixture(name), name)
    const before = await parts(fixture(name))
    const after = await parts(await saveXlsx(read.workbook, { source: read.source }))
    for (const [part, xml] of before)
      if (/^xl\/pivot(Cache|Tables)\//.test(part)) expect(after.get(part), part).toBe(xml)
    expect(after.get("xl/workbook.xml")).toMatch(
      /<pivotCaches><pivotCache cacheId="1\d" r:id="rId\d+"\/><\/pivotCaches>/,
    )
  })

  it("moves a PivotTable with rows inserted above it, patching only its location", async () => {
    const read = await openXlsx(fixture("pivot-basic.xlsx"), "pivot-basic.xlsx")
    const wb = read.workbook
    const index = wb.sheets.findIndex((s) => s.name === "Cross")
    wb.transact([{ kind: "insertRows", sheet: index, at: 0, count: 2 }], { label: "Insert" })
    expect(wb.sheets[index].pivots[0].location).toMatchObject({ r0: 6, r1: 10 })
    const after = await parts(await saveXlsx(wb, { source: read.source }))
    const xml = [...after.entries()].find(([, x]) => x.includes('name="PTCross"'))![1]
    expect(xml).toContain('<location ref="A7:D11"')
    const again = await openXlsx(await saveXlsx(wb, { source: read.source }))
    const cross = again.workbook.sheets[index].pivots[0]
    expect(cross.location).toEqual({ r0: 6, c0: 0, r1: 10, c1: 3 })
    expect(expectCellsAsExcel(again.workbook)).toBe(6)
  })

  it.each(FILES)("writes %s's caches and PivotTables anew, and they read back the same", async (name) => {
    const read = await openXlsx(fixture(name), name)
    const wb = read.workbook
    const tablesBefore = wb.sheets.map((s) => s.pivots.map(pivotTableKey))
    const cacheBefore = pivotCacheKey(wb.pivotCaches[0])
    // A refreshed cache (new records array) makes Cells write everything anew.
    wb.pivotCaches[0].records = wb.pivotCaches[0].records.map((r) => [...r])
    const saved = await saveXlsx(wb, { source: read.source })
    const after = await parts(saved)
    const def = [...after.entries()].find(([p]) => p.startsWith("xl/pivotCache/pivotCacheDefinition"))![1]
    expect(def).not.toContain("refreshedBy")
    const again = (await openXlsx(saved, name)).workbook
    expect(again.sheets.map((s) => s.pivots.map(pivotTableKey))).toEqual(tablesBefore)
    expect(pivotCacheKey(again.pivotCaches[0])).toBe(cacheBefore)
    expect(again.pivotCaches[0].records).toEqual(read.workbook.pivotCaches[0].records)
    expect(expectCellsAsExcel(again)).toBe(6)
  })

  it("writes a changed PivotTable anew and keeps the others as read", async () => {
    const read = await openXlsx(fixture("pivot-basic.xlsx"), "pivot-basic.xlsx")
    const wb = read.workbook
    const sheet = wb.sheets.find((s) => s.name === "Compact")!
    const changed = { ...sheet.pivots[0], rowGrandTotals: false }
    sheet.pivots = [changed]
    const before = await parts(fixture("pivot-basic.xlsx"))
    const after = await parts(await saveXlsx(wb, { source: read.source }))
    expect(after.get("xl/pivotTables/pivotTable1.xml")).not.toBe(before.get("xl/pivotTables/pivotTable1.xml"))
    expect(after.get("xl/pivotTables/pivotTable1.xml")).toContain('rowGrandTotals="0"')
    expect(after.get("xl/pivotTables/pivotTable2.xml")).toBe(before.get("xl/pivotTables/pivotTable2.xml"))
    expect(after.get("xl/pivotCache/pivotCacheRecords1.xml")).toBe(before.get("xl/pivotCache/pivotCacheRecords1.xml"))
    // The rows Excel had for the grand total line are gone from the written items.
    expect(after.get("xl/pivotTables/pivotTable1.xml")).not.toContain('t="grand"')
  })

  it("writes a PivotTable made in Cells", async () => {
    const wb = new Workbook(["Data", "Report"])
    const data = wb.sheets[0]
    const rows: (string | number)[][] = [
      ["Region", "Sales"],
      ["North", 10],
      ["South", 20],
      ["North", 5],
    ]
    rows.forEach((r, row) => r.forEach((value, col) => data.cells.set(row, col, { value })))
    const cache: PivotCache = {
      id: 1,
      source: { kind: "range", sheet: "Data", ref: "A1:B4" },
      fields: [
        { name: "Region", items: null },
        { name: "Sales", items: null },
      ],
      records: rows.slice(1),
      refreshOnLoad: false,
    }
    const fields = [newPivotField(0), newPivotField(0)]
    fields[0].axis = "row"
    fields[1].dataField = true
    const table: PivotTable = {
      name: "PivotTable1",
      cacheId: 1,
      at: { row: 2, col: 0 },
      location: { r0: 2, c0: 0, r1: 5, c1: 1 },
      fields,
      rows: [0],
      cols: [],
      pages: [],
      values: [{ name: "Sum of Sales", field: 1, fn: "sum", numFmt: "#,##0.00" }],
      rowGrandTotals: true,
      colGrandTotals: true,
      compact: true,
      outline: true,
      compactData: true,
      captions: { data: "Values" },
      showError: false,
      showMissing: true,
      style: {
        name: "PivotStyleLight16",
        rowHeaders: true,
        colHeaders: true,
        rowStripes: false,
        colStripes: false,
        lastColumn: true,
      },
    }
    wb.pivotCaches.push(cache)
    const report: Sheet = wb.sheets[1]
    report.pivots.push(table)
    for (const c of pivotLayout(table, cache).cells)
      if (c.value !== undefined)
        applyOp(wb.sheets, {
          kind: "setCells",
          sheet: 1,
          cells: [{ row: c.row, col: c.col, data: { value: c.value } }],
        })
    const saved = await saveXlsx(wb)
    const p = await parts(saved)
    expect(p.get("xl/workbook.xml")).toContain('<pivotCaches><pivotCache cacheId="1" r:id=')
    expect(p.get("xl/pivotCache/pivotCacheDefinition1.xml")).toContain(
      '<sharedItems count="2"><s v="North"/><s v="South"/>',
    )
    expect(p.get("xl/pivotCache/pivotCacheRecords1.xml")).toContain('<r><x v="0"/><n v="10"/></r><r><x v="1"/>')
    expect(p.get("xl/pivotTables/pivotTable1.xml")).toContain(
      '<location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>',
    )
    expect(p.get("xl/pivotTables/pivotTable1.xml")).toMatch(
      /<dataField name="Sum of Sales" fld="1" baseField="0" baseItem="0" numFmtId="4"\/>/,
    )
    const again = (await openXlsx(saved)).workbook
    expect(again.sheets[1].pivots[0].values[0].numFmt).toBe("#,##0.00")
    expect(again.pivotCaches[0].fields[0].items).toEqual(["North", "South"])
    expect(expectCellsAsExcel(again)).toBe(1)
  })
})
