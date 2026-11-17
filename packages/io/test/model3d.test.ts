import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { model3dXml, Workbook, type SheetPicture } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)))
const parts = async (bytes: Uint8Array) => {
  const zip = await readZip(bytes)
  const text = (end: string) => new TextDecoder().decode(zip.find((e) => e.normalizedName.endsWith(end))?.data)
  return { zip, text }
}

describe("Insert › 3D Models in xlsx", () => {
  it("opens Excel's 3D model: its file, its view and the picture Excel made of it", async () => {
    const opened = await openXlsx(fixture("excel/model3d-turned.xlsx"), "model3d-turned.xlsx")
    expect(opened.warnings).toEqual([])
    const pic = opened.workbook.sheets[0]!.objects[0] as SheetPicture
    expect(pic.kind).toBe("picture")
    expect(pic.name).toBe("3D Model 1")
    expect(pic.src.startsWith("data:image/png;base64,")).toBe(true)
    const m = pic.model3d!
    expect(m.rot).toEqual({ x: 30, y: 45, z: 10 })
    expect(m.viewport).toBeCloseTo(288, 0)
    // The .glb as Excel keeps it ("glTF", version 2).
    expect(new TextDecoder().decode(m.data.subarray(0, 4))).toBe("glTF")
    expect(m.xml).toContain('<am3d:preTrans dx="-90000000" dy="0" dz="0"/>')
  })

  it("opens a model Excel left unturned (an empty am3d:rot) facing front", async () => {
    const opened = await openXlsx(fixture("excel/model3d.xlsx"), "model3d.xlsx")
    const m = (opened.workbook.sheets[0]!.objects[0] as SheetPicture).model3d!
    expect(m.rot).toEqual({ x: 0, y: 0, z: 0 })
    expect(m.viewport).toBeCloseTo(389.5, 0)
  })

  it("saves an unchanged model back as Excel wrote it", async () => {
    const opened = await openXlsx(fixture("excel/model3d-turned.xlsx"), "model3d-turned.xlsx")
    const before = await parts(fixture("excel/model3d-turned.xlsx"))
    const after = await parts(await saveXlsx(opened.workbook, { source: opened.source }))
    const anchor = (x: string) => /<mc:AlternateContent[\s\S]*<\/mc:AlternateContent>/.exec(x)![0]
    expect(anchor(after.text("drawings/drawing1.xml"))).toBe(anchor(before.text("drawings/drawing1.xml")))
    expect(after.text("drawings/_rels/drawing1.xml.rels")).toContain("relationships/model3d")
    expect(after.zip.some((e) => e.normalizedName.endsWith(".glb"))).toBe(true)
  })

  it("writes a turned model's new view and picture, keeping Office's camera and lights", async () => {
    const opened = await openXlsx(fixture("excel/model3d-turned.xlsx"), "model3d-turned.xlsx")
    const pic = opened.workbook.sheets[0]!.objects[0] as SheetPicture
    pic.model3d = { ...pic.model3d!, rot: { x: 0, y: 90, z: 0 } }
    pic.src =
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
    const { text, zip } = await parts(await saveXlsx(opened.workbook, { source: opened.source }))
    const xml = text("drawings/drawing1.xml")
    expect(xml).toContain('<am3d:rot ay="5400000"/>')
    expect(xml).toContain('<am3d:pos x="0" y="0" z="53886815"/>')
    expect(xml).toContain("<am3d:ptLight")
    // Both sides of the AlternateContent carry the same shape id.
    const ids = [...xml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map((m) => m[1])
    expect(ids).toHaveLength(2)
    expect(ids[0]).toBe(ids[1])
    const rels = text("drawings/_rels/drawing1.xml.rels")
    const modelRid = /<am3d:model3d\b[^>]*r:embed="(rId\d+)"/.exec(xml)![1]
    const blipRid = /<am3d:blip r:embed="(rId\d+)"/.exec(xml)![1]
    expect(rels).toMatch(new RegExp(`Id="${modelRid}" Type="[^"]*relationships/model3d"`))
    expect(rels).toMatch(new RegExp(`Id="${blipRid}" Type="[^"]*relationships/image"`))
    expect(xml).toContain(
      `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${blipRid}"/>`,
    )
    expect(zip.filter((e) => e.normalizedName.endsWith(".glb"))).toHaveLength(1)
  })

  it("writes a new model as Excel does: the .glb as model/gltf.binary, and the view Office starts with", async () => {
    const wb = new Workbook()
    const glb = fixture("excel/model3d-turned.xlsx") // any bytes stand for the model here
    wb.sheets[0]!.objects.push({
      kind: "picture",
      id: 1,
      name: "3D Model 1",
      anchor: { from: { col: 1, row: 1, x: 0, y: 0 }, to: { col: 4, row: 10, x: 0, y: 0 } },
      src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
      model3d: {
        data: glb,
        rot: { x: 0, y: 0, z: 0 },
        viewport: 288,
        xml: model3dXml({ meterPerUnit: 0.5, preTrans: [-2.5, 0, 0], cameraZ: 1.4968 }),
      },
    } as SheetPicture)
    const { text } = await parts(await saveXlsx(wb))
    expect(text("[Content_Types].xml")).toMatch(
      /model3d1\.glb" ContentType="model\/gltf\.binary"|Extension="glb" ContentType="model\/gltf\.binary"/,
    )
    const xml = text("drawings/drawing1.xml")
    expect(xml).toContain('Requires="am3d"')
    expect(xml).toContain('<am3d:meterPerModelUnit n="500000" d="1000000"/>')
    expect(xml).toContain('<am3d:preTrans dx="-90000000" dy="0" dz="0"/>')
    expect(xml).toContain('<am3d:objViewport viewportSz="2743200"/>')
    expect(xml).toContain("<am3d:rot/>")
    const reopened = await openXlsx(await saveXlsx(wb), "new.xlsx")
    expect((reopened.workbook.sheets[0]!.objects[0] as SheetPicture).model3d?.data.length).toBe(glb.length)
  })
})
