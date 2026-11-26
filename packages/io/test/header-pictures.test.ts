// Header and footer pictures, as Excel 365 saved them (fixtures/excel/header-picture.xlsx): a 200 × 100 px picture in
// the centre header as it is, and in the left footer (after "Logo ") at 90 × 30 pt, unlocked, 10 pt cropped from the
// top, grayscale, brightness 70%, contrast 40%.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { newHeaderPicture, Workbook } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/header-picture.xlsx", import.meta.url)))
const files = async (bytes: Uint8Array) => new Map((await readZip(bytes)).map((e) => [e.normalizedName, e.data]))
const text = (data: Uint8Array | undefined) => new TextDecoder().decode(data)

describe("header and footer pictures in xlsx", () => {
  it("reads Excel's, with Format Picture's settings", async () => {
    const { workbook } = await openXlsx(fixture(), "header-picture.xlsx")
    const { pictures, header, footer } = workbook.sheets[0].pageSetup
    expect(header.center).toBe("&G")
    expect(footer.left).toBe("Logo &G")
    expect(Object.keys(pictures).sort()).toEqual(["CH", "LF"])
    expect(pictures.CH).toMatchObject({
      title: "logo",
      width: 150,
      height: 75,
      crop: { top: 0, bottom: 0, left: 0, right: 0 },
      lockAspect: true,
      color: "automatic",
      brightness: 0.5,
      contrast: 0.5,
    })
    expect(pictures.CH.src).toMatch(/^data:image\/png;base64,/)
    const lf = pictures.LF
    expect(lf).toMatchObject({ width: 90, height: 26.25, lockAspect: false, color: "grayscale" })
    expect(lf.crop.top).toBeCloseTo(10 / 75, 3)
    expect(lf.brightness).toBeCloseTo(0.7, 3)
    expect(lf.contrast).toBeCloseTo(0.4, 3)
  })

  it("writes them as Excel does, one picture part for both, and reads them back", async () => {
    const { workbook, source } = await openXlsx(fixture(), "header-picture.xlsx")
    const out = await files(await saveXlsx(workbook, { source, fileName: "h.xlsx" }))
    const sheet = text(out.get("xl/worksheets/sheet1.xml"))
    expect(sheet.match(/<legacyDrawingHF r:id="(rId\d+)"\/>/g)).toHaveLength(1)
    const vmlName = [...out.keys()].find((k) => k.startsWith("xl/drawings/vmlDrawing"))!
    const vml = text(out.get(vmlName))
    expect(vml).toContain(`<v:shape id="CH"`)
    expect(vml).toContain(`width:150pt;height:75pt`)
    expect(vml).toMatch(/<v:shape id="LF"[^>]*width:90pt;height:26.25pt/)
    expect(vml).toContain(`croptop="8738f"`)
    expect(vml).toContain(`gain="52428f" blacklevel="13107f" grayscale="t"`)
    expect(vml).toContain(`aspectratio="f"`)
    expect([...out.keys()].filter((k) => k.startsWith("xl/media/"))).toHaveLength(1)
    const back = (await openXlsx(await saveXlsx(workbook, { source, fileName: "h.xlsx" }), "h.xlsx")).workbook
    expect(back.sheets[0].pageSetup.pictures.LF).toMatchObject({ width: 90, height: 26.25, color: "grayscale" })
  })

  it("a new workbook's picture in the first page's right footer, and none once taken out", async () => {
    const wb = new Workbook()
    const png =
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYPj/HwADAgH/eL9GtQAAAABJRU5ErkJggg=="
    const setup = wb.sheets[0].pageSetup
    setup.differentFirst = true
    setup.firstFooter = { left: "", center: "", right: "&G" }
    setup.pictures.RFFIRST = { ...newHeaderPicture(png, "dot", 96, 48), color: "blackAndWhite" }
    const out = await files(await saveXlsx(wb, { fileName: "n.xlsx" }))
    const vml = text(out.get([...out.keys()].find((k) => k.startsWith("xl/drawings/vmlDrawing"))!))
    expect(vml).toMatch(/<v:shape id="RFFIRST"[^>]*width:72pt;height:36pt/)
    expect(vml).toContain(`grayscale="t" bilevel="t"`)
    const back = (await openXlsx(await saveXlsx(wb, { fileName: "n.xlsx" }), "n.xlsx")).workbook.sheets[0]
    expect(back.pageSetup.pictures.RFFIRST).toMatchObject({ width: 72, height: 36, color: "blackAndWhite" })

    const { workbook, source } = await openXlsx(fixture(), "header-picture.xlsx")
    workbook.sheets[0].pageSetup.pictures = {}
    const cleared = await files(await saveXlsx(workbook, { source, fileName: "c.xlsx" }))
    expect(text(cleared.get("xl/worksheets/sheet1.xml"))).not.toContain("legacyDrawingHF")
    expect([...cleared.keys()].filter((k) => k.startsWith("xl/drawings/") || k.startsWith("xl/media/"))).toEqual([])
  })
})
