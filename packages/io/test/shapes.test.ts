import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import {
  defaultShape,
  isChart,
  type SheetGroup,
  type SheetObject,
  type SheetPicture,
  type SheetShape,
  Workbook,
} from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures/excel", `${name}.xlsx`)))
const byName = (list: SheetObject[], name: string) => list.find((o) => o.name === name)!
const drawingXml = async (bytes: Uint8Array) => {
  const entries = await readZip(bytes)
  const e = entries.find((x) => /drawings\/drawing\d+\.xml$/.test(x.name))
  return e ? new TextDecoder().decode(e.data) : ""
}

describe("Excel's shapes.xlsx (shapes, text box, connector, WordArt, group, picture)", () => {
  it("reads each object in order, with its look and text", async () => {
    const read = await openXlsx(fixture("shapes"), "shapes.xlsx")
    const objects = read.workbook.sheets[0].objects
    expect(objects.map((o) => (isChart(o) ? "chart" : o.kind))).toEqual([
      "shape",
      "shape",
      "shape",
      "shape",
      "shape",
      "shape",
      "group",
      "picture",
    ])
    const box = byName(objects, "Box") as SheetShape
    expect(box.geometry).toBe("rect")
    expect(box.fill).toEqual({ kind: "solid", color: { theme: "accent1" } })
    expect(box.line.width).toBe(1)
    expect(box.text?.paragraphs[0].runs[0]).toMatchObject({ text: "Hello shapes", size: 11, color: { theme: "lt1" } })
    expect(box.anchor.from).toMatchObject({ col: 2, row: 2 })

    const oval = objects[1] as SheetShape
    expect(oval).toMatchObject({ geometry: "ellipse", rotation: 30, fill: { kind: "solid", color: { rgb: "FF0000" } } })
    const textBox = objects[3] as SheetShape
    expect(textBox.textBox).toBe(true)
    expect(textBox.text?.paragraphs[0].runs[0]).toMatchObject({ text: "A text box", bold: true })
    const arrow = objects[4] as SheetShape
    expect(arrow).toMatchObject({ connector: true, geometry: "straightConnector1", line: { tail: "triangle" } })
    const wordArt = objects[5] as SheetShape
    expect(wordArt.wordArt).toBeDefined()
    expect(wordArt.anchor.to.col).toBeGreaterThan(wordArt.anchor.from.col)

    const group = byName(objects, "Pair") as SheetGroup
    expect(group.items.map((i) => (i as SheetShape).geometry)).toEqual(["roundRect", "triangle"])
    expect(group.items[1].box!.x).toBeGreaterThan(group.items[0].box!.x)

    const picture = objects[7] as SheetPicture
    expect(picture.src.startsWith("data:image/png;base64,")).toBe(true)
    expect(picture.altText).toBe("A green dot")
  })

  it("writes unchanged objects back as Excel wrote them, and changed ones anew", async () => {
    const read = await openXlsx(fixture("shapes"), "shapes.xlsx")
    const before = await drawingXml(fixture("shapes"))
    const box = byName(read.workbook.sheets[0].objects, "Box") as SheetShape
    box.text!.paragraphs[0].runs[0].text = "Changed"
    const saved = await saveXlsx(read.workbook, { source: read.source, views: read.views })
    const after = await drawingXml(saved)
    // Office's creation ids stay on the shapes Cells didn't change.
    expect(after.match(/a16:creationId/g)?.length).toBe(before.match(/a16:creationId/g)!.length - 1)
    expect(after).toContain("Changed")
    const again = await openXlsx(saved, "x.xlsx")
    const objects = again.workbook.sheets[0].objects
    expect(objects).toHaveLength(8)
    expect(((objects[0] as SheetShape).text!.paragraphs[0].runs[0] as { text: string }).text).toBe("Changed")
    expect((objects[7] as SheetPicture).src).toBe((read.workbook.sheets[0].objects[7] as SheetPicture).src)
  })

  it("writes shapes and pictures Cells made and reads them back the same", async () => {
    const wb = new Workbook()
    const sheet = wb.sheets[0]
    const anchor = { from: { row: 1, col: 1, x: 0, y: 0 }, to: { row: 5, col: 4, x: 10, y: 0 } }
    const shape: SheetShape = {
      kind: "shape",
      id: 1,
      name: "Oval 1",
      anchor,
      geometry: "ellipse",
      ...defaultShape("ellipse"),
      text: { paragraphs: [{ align: "ctr", runs: [{ text: "Hi & bye", bold: true, size: 14 }] }], anchor: "ctr" },
      rotation: 45,
    }
    const picture: SheetPicture = {
      kind: "picture",
      id: 2,
      name: "Picture 2",
      anchor: { from: { row: 6, col: 1, x: 0, y: 0 }, to: { row: 8, col: 2, x: 0, y: 0 } },
      src: "data:image/png;base64," + Buffer.from(fixture("shapes").subarray(0, 0)).toString("base64") + "iVBORw0KGgo=",
      placement: "move",
      altText: "Logo",
      crop: { l: 0.1, t: 0, r: 0.25, b: -0.05 },
      geometry: "heart",
    }
    sheet.objects.push(shape, picture)
    const again = await openXlsx(await saveXlsx(wb), "x.xlsx")
    const [s, p] = again.workbook.sheets[0].objects as [SheetShape, SheetPicture]
    expect(s).toMatchObject({ geometry: "ellipse", rotation: 45, fill: shape.fill, line: shape.line, anchor })
    expect(s.text!.paragraphs[0]).toEqual(shape.text!.paragraphs[0])
    // Picture Format › Crop: the crop (a negative side is Fit's empty space) and Crop to Shape.
    expect(p).toMatchObject({
      src: picture.src,
      placement: "move",
      altText: "Logo",
      crop: picture.crop,
      geometry: "heart",
    })
  })
})

describe("the workbook's theme", () => {
  it("reads the theme's name, colours and fonts", async () => {
    const read = await openXlsx(fixture("shapes"), "shapes.xlsx")
    expect(read.workbook.theme.name).toBe("Office Theme")
    expect(read.workbook.theme.colors.accent1).toMatch(/^[0-9A-F]{6}$/)
    expect(read.workbook.theme.fonts.minor).toBeTruthy()
  })
})

describe("Page Layout › Background", () => {
  const PNG =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYPj/HwADAgH/eL9GtQAAAABJRU5ErkJggg=="
  it("writes the sheet's picture and reads it back; a deleted one leaves no part behind", async () => {
    const wb = new Workbook(["One", "Two"])
    wb.sheets[1].background = PNG
    const bytes = await saveXlsx(wb)
    const entries = await readZip(bytes)
    const sheet2 = new TextDecoder().decode(entries.find((e) => e.name === "xl/worksheets/sheet2.xml")!.data)
    expect(sheet2).toMatch(/<picture r:id="rId\d+"\/><\/worksheet>$/)
    expect(entries.filter((e) => e.name.startsWith("xl/media/"))).toHaveLength(1)
    const again = await openXlsx(bytes, "x.xlsx")
    expect(again.workbook.sheets[0].background).toBeUndefined()
    expect(again.workbook.sheets[1].background).toBe(PNG)

    // Saved again unchanged: one picture still; deleted: none.
    const twice = await readZip(await saveXlsx(again.workbook, { source: again.source }))
    expect(twice.filter((e) => e.name.startsWith("xl/media/"))).toHaveLength(1)
    again.workbook.sheets[1].background = undefined
    const none = await readZip(await saveXlsx(again.workbook, { source: again.source }))
    expect(none.filter((e) => e.name.startsWith("xl/media/"))).toHaveLength(0)
    const rels = new TextDecoder().decode(none.find((e) => e.name === "xl/worksheets/_rels/sheet2.xml.rels")?.data)
    expect(rels).not.toMatch(/image/)
  })

  it("reads Excel's own background and saves it back where Excel keeps it", async () => {
    const read = await openXlsx(fixture("background"), "background.xlsx")
    const [report, notes] = read.workbook.sheets
    expect(report.background).toMatch(/^data:image\/png;base64,iVBOR/)
    expect(notes.background).toBeUndefined()
    const entries = await readZip(await saveXlsx(read.workbook, { source: read.source }))
    const sheet1 = new TextDecoder().decode(entries.find((e) => e.name === "xl/worksheets/sheet1.xml")!.data)
    expect(sheet1.match(/<picture r:id="(rId\d+)"\/>/g)).toHaveLength(1)
    expect(entries.filter((e) => e.name.startsWith("xl/media/"))).toHaveLength(1)
  })
})
