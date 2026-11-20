// Timelines on PivotTables that Excel 365 made (fixtures/excel/pivot-timeline.xlsx): "Date" filters PTOne to March –
// June 2025 at the months level; "Date 1" (caption "Order date", Dark 1, days level) is on PTTwo with nothing picked.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import {
  isChart,
  pivotLayout,
  pivotSlicerTarget,
  timelineSelection,
  withTimelineDates,
  type SheetTimeline,
  type Workbook,
} from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/pivot-timeline.xlsx", import.meta.url)))
async function parts(bytes: Uint8Array) {
  const zip = await readZip(bytes)
  return new Map(zip.map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
}
const timelines = (wb: Workbook) =>
  wb.sheets.flatMap((s) => s.objects.filter((o): o is SheetTimeline => !isChart(o) && o.kind === "timeline"))
const total = (wb: Workbook, name: string) => {
  const table = wb.sheets.flatMap((s) => s.pivots).find((t) => t.name === name)!
  const cells = pivotLayout(
    table,
    wb.pivotCaches.find((c) => c.id === table.cacheId)!,
  ).cells
  const label = cells.find((c) => c.value === "Grand Total")!
  return cells.find((c) => c.row === label.row && typeof c.value === "number")!.value
}
const picked = (wb: Workbook, o: SheetTimeline) => timelineSelection(pivotSlicerTarget(wb.sheets, wb.pivotCaches, o)!)

describe("timelines", () => {
  it("reads Excel's timelines: their PivotTables, settings and dates, and the PivotTables filtered by them", async () => {
    const { workbook: wb, warnings } = await openXlsx(fixture(), "pivot-timeline.xlsx")
    expect(warnings).toEqual([])
    const [date, order] = timelines(wb)
    expect(date).toMatchObject({ name: "Date", caption: "Date", level: "months", selectionLevel: "months" })
    expect(date.pivot).toMatchObject({ field: "Date", tables: [{ sheet: "One", name: "PTOne" }] })
    expect(order).toMatchObject({ caption: "Order date", level: "days", style: "TimeSlicerStyleDark1", header: true })
    // 2025-03-01 to 2025-06-30.
    expect(picked(wb, date)).toEqual({ from: 45717, to: 45838 })
    expect(picked(wb, order)).toBeNull()
    expect(total(wb, "PTOne")).toBe(61)
    expect(total(wb, "PTTwo")).toBe(6630)
  })

  it("writes them back as Excel does, a changed pick in the PivotTables and the timeline cache", async () => {
    const read = await openXlsx(fixture(), "pivot-timeline.xlsx")
    const wb = read.workbook
    // April only on PTTwo's timeline, as a click does it.
    const two = wb.sheets[2]
    two.pivots = two.pivots.map((t) => withTimelineDates(t, 2, { from: 45748, to: 45777 }))
    const bytes = await saveXlsx(wb, { source: read.source, fileName: "pivot-timeline.xlsx" })
    const out = await parts(bytes)
    const wbXml = out.get("xl/workbook.xml")!
    expect(wbXml).toContain('<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"')
    expect(wbXml).toContain('<definedName name="NativeTimeline_Date1">#N/A</definedName>')
    const caches = [...out].filter(([k]) => k.startsWith("xl/timelineCaches/")).map(([, v]) => v)
    expect(caches).toHaveLength(2)
    const second = caches.find((c) => c.includes('name="NativeTimeline_Date1"'))!
    expect(second).toContain('<pivotTables><pivotTable tabId="3" name="PTTwo"/></pivotTables>')
    expect(second).toContain(
      '<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="89352520" filterType="dateBetween">' +
        '<selection startDate="2025-04-01T00:00:00" endDate="2025-04-30T00:00:00"/>' +
        '<bounds startDate="2025-01-01T00:00:00" endDate="2026-01-01T00:00:00"/></state>',
    )
    const pt = [...out].find(([k, v]) => k.startsWith("xl/pivotTables/") && v.includes('name="PTTwo"'))![1]
    expect(pt).toContain('<filter fld="2" type="dateBetween" evalOrder="-1" id="1">')
    expect(pt).toContain('<x15:pivotFilter useWholeDay="1"/>')

    const { workbook: again, warnings } = await openXlsx(bytes, "pivot-timeline.xlsx")
    expect(warnings).toEqual([])
    const [date, order] = timelines(again)
    expect(picked(again, date)).toEqual({ from: 45717, to: 45838 })
    expect(picked(again, order)).toEqual({ from: 45748, to: 45777 })
    expect(order).toMatchObject({ caption: "Order date", level: "days", style: "TimeSlicerStyleDark1" })
    expect(total(again, "PTOne")).toBe(61)
    // Excel's total for April on PTTwo, after it refreshes Cells' file.
    expect(total(again, "PTTwo")).toBe(295)
  })
})
