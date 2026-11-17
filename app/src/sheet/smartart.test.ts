import { readFileSync } from "node:fs"
import { join } from "node:path"
import { OFFICE_THEME, type SheetGroup } from "@cells/core"
import { openXlsx } from "@cells/io"
import { describe, expect, it } from "vitest"
import { objectSvg } from "@/drawing/render"
import { SheetController } from "./controller"

const fixture = (name: string) =>
  new Uint8Array(readFileSync(join(__dirname, "../../../packages/io/test/fixtures", name)))

describe("SmartArt", () => {
  it("draws Excel's SmartArt as Office laid it out, and keeps it whole on Ungroup", async () => {
    const opened = await openXlsx(fixture("excel/smartart.xlsx"), "smartart.xlsx")
    const c = new SheetController(opened.workbook)
    const g = c.sheet.objects[0] as SheetGroup
    const box = c.objects.box(g)
    const svg = objectSvg(g, box.w, box.h, OFFICE_THEME)
    expect(svg).toContain(">Plan<")
    // Five blocks in the theme's first accent.
    expect(svg.match(/fill="#156082"/gi)?.length).toBeGreaterThanOrEqual(5)
    c.objects.select(g.id)
    c.objects.ungroup()
    expect(c.sheet.objects).toHaveLength(1)
    expect((c.sheet.objects[0] as SheetGroup).smartArt).toBeDefined()
  })
})
