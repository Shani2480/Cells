import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { Workbook } from "@cells/core"
import { DEFAULT_VIEW, openXlsx, saveXlsx } from "../src"

const text = (files: { name: string; data: Uint8Array }[], name: string) =>
  new TextDecoder().decode(files.find((f) => f.name === name)!.data)

describe("a workbook's windows (View › New Window)", () => {
  it("writes a workbookView per window and a sheetView per window in each sheet, and reads them back", async () => {
    const wb = new Workbook(["Budget", "Trips"])
    const bytes = await saveXlsx(wb, {
      views: [{ ...DEFAULT_VIEW, zoom: 1.5 }, { ...DEFAULT_VIEW }],
      activeSheet: 0,
      windows: [
        {
          activeSheet: 1,
          views: [
            { ...DEFAULT_VIEW },
            { ...DEFAULT_VIEW, zoom: 0.75, active: { row: 4, col: 2 }, selection: [{ r0: 4, c0: 2, r1: 4, c1: 2 }] },
          ],
        },
      ],
    })
    const files = await readZip(bytes)
    expect(text(files, "xl/workbook.xml")).toMatch(
      /<bookViews><workbookView\/><workbookView activeTab="1"\/><\/bookViews>/,
    )
    const trips = text(files, "xl/worksheets/sheet2.xml")
    expect(trips).toMatch(/<sheetView workbookViewId="0"\/>/)
    expect(trips).toMatch(
      /<sheetView tabSelected="1" zoomScale="75" zoomScaleNormal="75" workbookViewId="1"><selection activeCell="C5" sqref="C5"\/>/,
    )

    const read = await openXlsx(bytes)
    expect(read.activeSheet).toBe(0)
    expect(read.views[0].zoom).toBe(1.5)
    expect(read.windows).toHaveLength(1)
    expect(read.windows[0].activeSheet).toBe(1)
    expect(read.windows[0].views[1].zoom).toBe(0.75)
    expect(read.windows[0].views[1].active).toEqual({ row: 4, col: 2 })

    // Closing the second window: saved again with no others, the file has one window.
    const again = await readZip(await saveXlsx(read.workbook, { source: read.source, views: read.views, windows: [] }))
    expect(text(again, "xl/workbook.xml")).not.toMatch(/activeTab="1"/)
    expect(text(again, "xl/worksheets/sheet2.xml")).not.toMatch(/workbookViewId="1"/)
    // Left out, the windows the file had are kept.
    const kept = await readZip(await saveXlsx(read.workbook, { source: read.source, views: read.views }))
    expect(text(kept, "xl/workbook.xml")).toMatch(/activeTab="1"/)
  })

  it("a workbook with one window reads no others", async () => {
    const read = await openXlsx(await saveXlsx(new Workbook()))
    expect(read.windows).toEqual([])
  })

  it("reads the windows of a file Excel saved with New Window, and keeps each window's size and place", async () => {
    const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/windows.xlsx", import.meta.url)))
    const read = await openXlsx(bytes)
    expect(read.windows).toHaveLength(1)
    expect(read.windows[0].activeSheet).toBe(1)
    expect(read.windows[0].views[1].zoom).toBe(0.7)
    expect(read.windows[0].views[1].active).toEqual({ row: 11, col: 4 })
    expect(read.views[1].zoom).toBe(1)
    const files = await readZip(
      await saveXlsx(read.workbook, { source: read.source, views: read.views, activeSheet: 0, windows: read.windows }),
    )
    const views = text(files, "xl/workbook.xml").match(/<bookViews>.*<\/bookViews>/)![0]
    expect(views.match(/<workbookView /g)).toHaveLength(2)
    expect(views).toMatch(/xr2:uid="\{9B8A21D2-CC06-4F7A-B0FE-75014A2564A1\}"/)
    expect(views).toMatch(/<workbookView activeTab="1" xWindow="-120"[^>]*windowWidth="15600"[^>]*\{9B8A21D2/)
  })
})
