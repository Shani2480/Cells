import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { copyObject, smartArtParts, Workbook, type SheetGroup, type SheetShape, type SmartArt } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)))
const parts = async (bytes: Uint8Array) => {
  const zip = await readZip(bytes)
  const text = (end: string) => new TextDecoder().decode(zip.find((e) => e.normalizedName.endsWith(end))?.data)
  return { zip, text }
}

describe("SmartArt in xlsx", () => {
  it("opens Excel's SmartArt as the shapes Office drew, with its text as a tree", async () => {
    const opened = await openXlsx(fixture("excel/smartart.xlsx"), "smartart.xlsx")
    expect(opened.warnings).toEqual([])
    const g = opened.workbook.sheets[0]!.objects[0] as SheetGroup
    expect(g.kind).toBe("group")
    expect(g.name).toBe("Diagram 1")
    expect(g.smartArt!.layout).toBe("urn:microsoft.com/office/officeart/2005/8/layout/default")
    expect(g.smartArt!.colors).toBe("urn:microsoft.com/office/officeart/2005/8/colors/accent1_2")
    expect(g.smartArt!.nodes.map((n) => n.text)).toEqual(["Plan", "", "", "", ""])
    expect(g.items).toHaveLength(5)
    const first = g.items[0] as SheetShape
    expect(first.geometry).toBe("rect")
    expect(first.text!.paragraphs[0]!.runs[0]!.text).toBe("Plan")
    // In the frame's own pixels.
    expect(first.box!.x).toBeCloseTo(572802 / 9525)
  })

  it("saves it back as Excel wrote it while it's unchanged, moved or not", async () => {
    const opened = await openXlsx(fixture("excel/smartart.xlsx"), "smartart.xlsx")
    const g = opened.workbook.sheets[0]!.objects[0] as SheetGroup
    g.anchor = { from: { ...g.anchor.from, row: 3 }, to: { ...g.anchor.to, row: 16 } }
    const before = await parts(fixture("excel/smartart.xlsx"))
    const after = await parts(await saveXlsx(opened.workbook, { source: opened.source }))
    const frame = (x: string) => /<xdr:graphicFrame[\s\S]*<\/xdr:graphicFrame>/.exec(x)![0]
    expect(frame(after.text("drawings/drawing1.xml"))).toBe(frame(before.text("drawings/drawing1.xml")))
    expect(after.text("drawings/drawing1.xml")).toContain("<xdr:row>3</xdr:row>")
    for (const p of ["diagrams/data1.xml", "diagrams/layout1.xml", "diagrams/drawing1.xml"])
      expect(after.text(p)).toBe(before.text(p))
  })

  it("writes a copy with parts of its own, Office's drawing named by its data", async () => {
    const opened = await openXlsx(fixture("excel/smartart.xlsx"), "smartart.xlsx")
    const sheet = opened.workbook.sheets[0]!
    const copy = copyObject(sheet.objects[0] as SheetGroup)
    copy.id = 99
    copy.name = "Diagram 2"
    delete copy.source
    sheet.objects.push(copy)
    const { text, zip } = await parts(await saveXlsx(opened.workbook, { source: opened.source }))
    const drawing = text("drawings/drawing1.xml")
    expect(drawing.match(/<dgm:relIds/g)).toHaveLength(2)
    const rels = text("drawings/_rels/drawing1.xml.rels")
    expect(rels.match(/relationships\/diagramData"/g)).toHaveLength(2)
    expect(zip.filter((e) => e.normalizedName.includes("diagrams/"))).toHaveLength(10)
    const data2 = zip.find((e) => /diagrams\/data2\.xml$/.test(e.normalizedName))!
    const relId = /dataModelExt\b[^>]*relId="([^"]+)"/.exec(new TextDecoder().decode(data2.data))![1]
    expect(rels).toMatch(new RegExp(`Id="${relId}" Type="[^"]*diagramDrawing" Target="\\.\\./diagrams/drawing2\\.xml"`))
    expect(text("[Content_Types].xml")).toContain('PartName="/xl/diagrams/layout2.xml"')
    // And it opens again as SmartArt.
    const again = await openXlsx(await saveXlsx(opened.workbook, { source: opened.source }), "again.xlsx")
    expect(again.workbook.sheets[0]!.objects.filter((o) => "kind" in o && o.kind === "group")).toHaveLength(2)
  })
})

describe("SmartArt made in Cells", () => {
  it("is written with Cells' layout and colours and no drawing, and opens with its text", async () => {
    const wb = new Workbook()
    const smartArt: SmartArt = {
      layout: "urn:cells.app/smartart/layout/process",
      nodes: [
        { text: "Plan", children: [{ text: "Scope", children: [] }] },
        { text: "Build", children: [] },
        { text: "", children: [] },
      ],
      colors: "urn:cells.app/smartart/colors/accent2",
    }
    smartArt.parts = smartArtParts(smartArt)
    wb.sheets[0]!.objects.push({
      kind: "group",
      id: 1,
      name: "Diagram 1",
      anchor: { from: { row: 1, col: 1, x: 0, y: 0 }, to: { row: 15, col: 8, x: 0, y: 0 } },
      child: { w: 480, h: 320 },
      items: [],
      smartArt,
    })
    const bytes = await saveXlsx(wb)
    const { zip, text } = await parts(bytes)
    expect(zip.some((e) => /diagrams\/drawing/.test(e.normalizedName))).toBe(false)
    expect(text("diagrams/layout1.xml")).toContain('uniqueId="urn:cells.app/smartart/layout/process"')
    expect(text("diagrams/colors1.xml")).toContain('uniqueId="urn:cells.app/smartart/colors/accent2"')
    const again = (await openXlsx(bytes, "made.xlsx")).workbook.sheets[0]!.objects[0] as SheetGroup
    expect(again.smartArt!.layout).toBe("urn:cells.app/smartart/layout/process")
    expect(again.smartArt!.colors).toBe("urn:cells.app/smartart/colors/accent2")
    expect(again.smartArt!.nodes).toEqual(smartArt.nodes)
  })
})
