// PivotTables in xlsx: the files Excel 365 made (fixtures/excel/pivot-basic.xlsx and pivot-more.xlsx; the same
// made-up sales data with PivotTables in compact, outline and tabular form, crossed, with report filters, every
// summary function, custom subtotals, blank rows, collapsed items, captions and sorting).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  applyOp,
  cellName,
  newPivotField,
  pivotCellLook,
  pivotLayout,
  pivotStyleDef,
  PIVOT_STYLE_GROUPS,
  rangeGroupItems,
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

  it("reads Label Filters, Value Filters and Top 10, and the engine shows Excel's cells (pivot-filters.xlsx)", async () => {
    const { workbook: wb } = await openXlsx(fixture("pivot-filters.xlsx"), "pivot-filters.xlsx")
    expect(expectCellsAsExcel(wb)).toBe(15)
  })

  it("reads Show Values As, and the engine shows Excel's cells (pivot-showas.xlsx)", async () => {
    const { workbook: wb } = await openXlsx(fixture("pivot-showas.xlsx"), "pivot-showas.xlsx")
    expect(expectCellsAsExcel(wb)).toBe(22)
    const all = wb.sheets.flatMap((s) => s.pivots)
    const value = (n: string) => all.find((t) => t.name === `PT${n}`)!.values[0]
    expect(value("PctOf")).toMatchObject({ showAs: "percent", baseField: 0, baseItem: 0, numFmt: "0.00%" })
    expect(value("PctDiffPrev")).toMatchObject({ showAs: "percentDiff", baseField: 1, baseItem: "previous" })
    expect(value("DiffNext").baseItem).toBe("next")
    expect(value("RankDesc")).toMatchObject({ showAs: "rankDescending", baseField: 1 })
    expect(value("PctTotal").baseField).toBeUndefined()
  })

  it("reads calculated fields, and the engine shows Excel's cells (pivot-calc.xlsx)", async () => {
    const { workbook: wb } = await openXlsx(fixture("pivot-calc.xlsx"), "pivot-calc.xlsx")
    expect(expectCellsAsExcel(wb)).toBe(3)
    const fields = wb.pivotCaches[0].fields
    expect(fields.slice(7).map((f) => [f.name, f.formula])).toEqual([
      ["Price", "Amount/Units"],
      ["Tax", "Amount*0.2"],
      ["Big", "IF(Units>10,Amount,0)"],
      ["Per Cost", "Amount/'Unit Cost'"],
    ])
    expect(wb.pivotCaches[0].records[0]).toHaveLength(7)
  })

  it("reads grouped fields, and the engine shows Excel's cells (pivot-group.xlsx)", async () => {
    const { workbook: wb } = await openXlsx(fixture("pivot-group.xlsx"), "pivot-group.xlsx")
    expect(wb.pivotCaches).toHaveLength(8)
    expect(expectCellsAsExcel(wb)).toBe(8)
    // Cells makes the same group items as Excel from the same settings.
    for (const cache of wb.pivotCaches)
      cache.fields.forEach((f) => {
        const g = f.group
        if (g?.kind !== "range") return
        const base = cache.records.map((r) => r[g.base] ?? null)
        const whole =
          Number.isInteger(g.start) &&
          Number.isInteger(g.interval) &&
          base.every((v) => typeof v !== "number" || Number.isInteger(v))
        const blank = base.some((v) => v === null || v === "")
        const made = rangeGroupItems(g.by, g.start, g.end, g.interval, g.by === "range" && whole, blank)
        expect(made, f.name).toEqual(g.items)
      })
    const text = wb.pivotCaches[6].fields[6]
    expect(text).toEqual({
      name: "Product2",
      items: null,
      group: { kind: "discrete", base: 1, map: [1, 0, 1], items: ["Plums", "Group1"] },
    })
  })

  it("writes grouped fields as Excel does, and reads them back the same", async () => {
    const read = await openXlsx(fixture("pivot-group.xlsx"), "pivot-group.xlsx")
    const wb = read.workbook
    for (const sheet of wb.sheets)
      sheet.pivots = sheet.pivots.map((t) => ({ ...t, style: { ...t.style, rowStripes: !t.style.rowStripes } }))
    wb.pivotCaches = wb.pivotCaches.map((c) => ({ ...c, refreshOnLoad: !c.refreshOnLoad }))
    const saved = await saveXlsx(wb, { source: read.source })
    const before = await parts(fixture("pivot-group.xlsx"))
    const after = await parts(saved)
    const groups = (files: Map<string, string>) =>
      [...files]
        .filter(([n]) => /pivotCacheDefinition\d+\.xml$/.test(n))
        .sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true }))
        .map(([, x]) => x.match(/<fieldGroup.*?(?:<\/fieldGroup>|\/>)/g))
    expect(groups(after)).toEqual(groups(before))
    const again = await openXlsx(saved)
    expect(again.workbook.pivotCaches.map((c) => c.fields)).toEqual(wb.pivotCaches.map((c) => c.fields))
    expect(expectCellsAsExcel(again.workbook)).toBe(8)
  })

  it("writes calculated fields as Excel does, and reads them back the same", async () => {
    const read = await openXlsx(fixture("pivot-calc.xlsx"), "pivot-calc.xlsx")
    const wb = read.workbook
    for (const sheet of wb.sheets)
      sheet.pivots = sheet.pivots.map((t) => ({ ...t, style: { ...t.style, rowStripes: !t.style.rowStripes } }))
    // The cache written anew too.
    wb.pivotCaches = wb.pivotCaches.map((c) => ({ ...c, refreshOnLoad: !c.refreshOnLoad }))
    const saved = await saveXlsx(wb, { source: read.source })
    const files = await parts(saved)
    const def = [...files].find(([n]) => /pivotCacheDefinition\d+\.xml$/.test(n))![1]
    expect(def).toContain('<cacheField name="Price" numFmtId="0" formula="Amount/Units" databaseField="0"/>')
    expect(def).toContain(`formula="Amount/'Unit Cost'"`)
    const records = [...files].find(([n]) => /pivotCacheRecords\d+\.xml$/.test(n))![1]
    expect(records.match(/<r>(.*?)<\/r>/)![1].match(/<[xnsmdbe] /g)).toHaveLength(7)
    const again = await openXlsx(saved)
    expect(again.workbook.pivotCaches[0].fields.map((f) => f.formula)).toEqual(
      wb.pivotCaches[0].fields.map((f) => f.formula),
    )
    expect(expectCellsAsExcel(again.workbook)).toBe(3)
  })

  it("writes Show Values As as Excel does, and reads it back the same", async () => {
    const read = await openXlsx(fixture("pivot-showas.xlsx"), "pivot-showas.xlsx")
    const wb = read.workbook
    const before = wb.sheets.flatMap((s) => s.pivots.map((t) => t.values))
    for (const sheet of wb.sheets)
      sheet.pivots = sheet.pivots.map((t) => ({ ...t, style: { ...t.style, rowStripes: !t.style.rowStripes } }))
    const saved = await saveXlsx(wb, { source: read.source })
    const xml = [...(await parts(saved))].filter(([n]) => /^xl\/pivotTables\/[^/]+\.xml$/.test(n)).map(([, x]) => x)
    const pick = (name: string) => xml.find((x) => x.includes(`name="PT${name}"`))!
    expect(pick("PctOfEmpty")).toMatch(/showDataAs="percent" baseField="0" baseItem="2"/)
    expect(pick("PctDiffPrev")).toContain('baseItem="1048828"')
    expect(pick("PctParent")).toContain('<x14:dataField pivotShowAs="percentOfParent"/>')
    expect(pick("PctParent")).not.toContain("showDataAs")
    const again = await openXlsx(saved)
    expect(again.workbook.sheets.flatMap((s) => s.pivots.map((t) => t.values))).toEqual(before)
    expect(expectCellsAsExcel(again.workbook)).toBe(22)
  })

  it("writes the filters anew as Excel does, and reads them back the same", async () => {
    const read = await openXlsx(fixture("pivot-filters.xlsx"), "pivot-filters.xlsx")
    const wb = read.workbook
    const before = wb.sheets.flatMap((s) => s.pivots.map((t) => t.filters))
    // Every PivotTable changed a little, so all are written anew.
    for (const sheet of wb.sheets)
      sheet.pivots = sheet.pivots.map((t) => ({ ...t, style: { ...t.style, rowStripes: !t.style.rowStripes } }))
    const saved = await saveXlsx(wb, { source: read.source })
    const xml = [...(await parts(saved))].filter(([n]) => /^xl\/pivotTables\/[^/]+\.xml$/.test(n)).map(([, x]) => x)
    expect(xml).toHaveLength(15)
    // All but AxisHidden (unticked items only) have filters.
    expect(xml.filter((x) => x.includes("<filters"))).toHaveLength(14)
    for (const x of xml.filter((x) => x.includes("<filters")))
      expect(x).toMatch(/<filters count="\d"><filter fld="\d" type="\w+" evalOrder="-1" id="\d"/)
    const again = (await openXlsx(saved)).workbook
    expect(again.sheets.flatMap((s) => s.pivots.map((t) => t.filters))).toEqual(before)
    expect(again.sheets.find((s) => s.name === "TwoValues")!.pivots[0].hideValuesRow).toBe(true)
    expect(expectCellsAsExcel(again)).toBe(15)
  })

  it("a renamed PivotTable keeps what Cells keeps unread (Excel 365's no-Values-row extension)", async () => {
    const read = await openXlsx(fixture("pivot-filters.xlsx"), "pivot-filters.xlsx")
    const sheet = read.workbook.sheets.find((s) => s.name === "TwoValues")!
    sheet.pivots = [{ ...sheet.pivots[0], name: "Sales by Region" }]
    const saved = await saveXlsx(read.workbook, { source: read.source })
    const xml = [...(await parts(saved))].find(([, x]) => x.includes('name="Sales by Region"'))![1]
    expect(xml).toContain('hideValuesRow="1"')
    expect(xml).toContain("pivotTableDefinition16")
    const again = (await openXlsx(saved)).workbook.sheets.find((s) => s.name === "TwoValues")!
    expect(again.pivots[0]).toMatchObject({ name: "Sales by Region", hideValuesRow: true })
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

  it("writes Show expand/collapse buttons off, and reads it back", async () => {
    const read = await openXlsx(fixture("pivot-basic.xlsx"), "pivot-basic.xlsx")
    const sheet = read.workbook.sheets.find((s) => s.name === "Compact")!
    expect(sheet.pivots[0].showDrill).toBeUndefined()
    sheet.pivots = [{ ...sheet.pivots[0], showDrill: false }]
    const saved = await saveXlsx(read.workbook, { source: read.source })
    expect((await parts(saved)).get("xl/pivotTables/pivotTable1.xml")).toContain('showDrill="0"')
    const again = (await openXlsx(saved)).workbook.sheets.find((s) => s.name === "Compact")!
    expect(again.pivots[0].showDrill).toBe(false)
  })

  it("reads and writes Autofit column widths on update (useAutoFormatting)", async () => {
    const read = await openXlsx(fixture("pivot-basic.xlsx"), "pivot-basic.xlsx")
    const sheet = read.workbook.sheets.find((s) => s.name === "Compact")!
    expect(sheet.pivots[0].autoFit).toBeUndefined()
    sheet.pivots = [{ ...sheet.pivots[0], autoFit: false }]
    const saved = await saveXlsx(read.workbook, { source: read.source })
    expect((await parts(saved)).get("xl/pivotTables/pivotTable1.xml")).not.toContain("useAutoFormatting")
    const again = (await openXlsx(saved)).workbook.sheets.find((s) => s.name === "Compact")!
    expect(again.pivots[0].autoFit).toBe(false)
  })

  it("writes Enable show details off, and reads it back", async () => {
    const read = await openXlsx(fixture("pivot-basic.xlsx"), "pivot-basic.xlsx")
    const sheet = read.workbook.sheets.find((s) => s.name === "Compact")!
    expect(sheet.pivots[0].enableDrill).toBeUndefined()
    sheet.pivots = [{ ...sheet.pivots[0], enableDrill: false }]
    const saved = await saveXlsx(read.workbook, { source: read.source })
    expect((await parts(saved)).get("xl/pivotTables/pivotTable1.xml")).toContain('enableDrill="0"')
    const again = (await openXlsx(saved)).workbook.sheets.find((s) => s.name === "Compact")!
    expect(again.pivots[0].enableDrill).toBe(false)
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

  it("draws each PivotTable in its style as Excel does", async () => {
    const { workbook: wb } = await openXlsx(fixture("pivot-styles.xlsx"), "pivot-styles.xlsx")
    // What Excel 365 drew for these cells (its pictures of the tables, sampled).
    const look = (sheet: string, row: number, col: number) => {
      const t = wb.sheets.find((s) => s.name === sheet)!.pivots[0]
      return pivotCellLook(t, cacheOf(wb, t), row, col, pivotStyleDef(t.style.name))
    }
    const light = (hex: string) => ({ style: "thin", color: hex })
    // Light16, compact: the header row, an outer item (a subheading: bold, a line below), an inner one, Grand Total.
    expect(look("Compact", 2, 0)).toMatchObject({ fill: "#C0E6F5", bold: true, borderBottom: light("#44B3E1") })
    expect(look("Compact", 3, 0)).toEqual({ bold: true, fontColor: "#000000", borderBottom: light("#44B3E1") })
    expect(look("Compact", 4, 0)).toEqual({ indent: 1 })
    expect(look("Compact", 19, 1)).toMatchObject({ fill: "#C0E6F5", bold: true, borderTop: light("#44B3E1") })
    // Dark8 with the values down the rows: outer items grey, the rest the table's fill, the totals dark.
    expect(look("ValuesRows", 4, 0)).toMatchObject({ fill: "#BFBFBF", bold: true })
    expect(look("ValuesRows", 5, 1)).toMatchObject({ fill: "#D9D9D9" })
    expect(look("ValuesRows", 16, 2)).toMatchObject({ fill: "#404040", fontColor: "#FFFFFF", bold: true })
    // Light20, tabular with banded rows: the outer labels are a subheading down their group, lined where it ends.
    expect(look("Tabular", 3, 0)).toEqual({ fill: "#D9D9D9", bold: true, fontColor: "#000000" })
    expect(look("Tabular", 5, 0)?.borderBottom).toEqual(light("#D86DCD"))
    expect(look("Tabular", 4, 0)?.borderBottom).toBeUndefined()
    expect(look("Tabular", 4, 1)).toEqual({})
    // Dark2: the report filter, and the Grand Total column with banded columns.
    expect(look("Cross", 2, 0)).toMatchObject({ fill: "#0B3040", bold: true, fontColor: "#FFFFFF" })
    expect(look("Cross", 6, 3)).toMatchObject({ fill: "#83CCEB" })
    // Outside the table there's no look; a table without a style only indents.
    expect(look("Compact", 0, 5)).toBeNull()
    expect(PIVOT_STYLE_GROUPS.flatMap((g) => g.names).filter((n) => pivotStyleDef(n))).toHaveLength(84)
  })
})
