import { describe, expect, it } from "vitest"
import { isChart, type SheetShape } from "@cells/core"
import { SheetController } from "./controller"
import { chartBox } from "./charts"

const PNG = "data:image/png;base64,iVBORw0KGgo="

function setup() {
  const c = new SheetController()
  return { c, o: c.objects }
}

const ids = (c: SheetController) => c.sheet.objects.map((o) => o.id)
const names = (c: SheetController) => c.sheet.objects.map((o) => o.name)

describe("inserting objects", () => {
  it("puts shapes, text boxes, pictures and WordArt on the sheet with Excel's names, one undo each", () => {
    const { c, o } = setup()
    const rect = o.insertShape("rect")
    const box = o.insertTextBox()
    const pic = o.insertPicture(PNG, { w: 40, h: 30 })
    const icon = o.insertPicture(PNG, { w: 96, h: 96 }, { icon: true })
    o.insertWordArt({ wordArt: { effects: { shadow: true } } })
    expect(names(c)).toEqual(["Rectangle 1", "TextBox 2", "Picture 3", "Graphic 4", "Rectangle 5"])
    expect(rect.fill).toEqual({ kind: "solid", color: { theme: "accent1" } })
    expect(box.textBox).toBe(true)
    expect(chartBox(c.sheet, pic.anchor)).toMatchObject({ w: 40, h: 30 })
    expect(icon.icon).toBe(true)
    expect(o.selected).toEqual([5])
    c.undo()
    c.undo()
    expect(names(c)).toEqual(["Rectangle 1", "TextBox 2", "Picture 3"])
    c.redo()
    expect(names(c)).toHaveLength(4)
  })

  it("draws a pending shape where the pointer dragged, and a text box opens for typing", () => {
    const { c, o } = setup()
    o.startShape("ellipse")
    expect(o.pending?.geometry).toBe("ellipse")
    const drawn = o.drawPending({ x: 100, y: 50, w: 80, h: 40 })!
    expect(o.pending).toBeNull()
    expect(chartBox(c.sheet, drawn.anchor)).toMatchObject({ x: 100, y: 50, w: 80, h: 40 })
    o.startShape("rect", true)
    const tb = o.drawPending({ x: 10, y: 10, w: 120, h: 30 })!
    expect(o.editing).toBe(tb.id)
    o.setText(tb.id, "Notes\nfor June")
    const text = (c.sheet.objects.at(-1) as SheetShape).text!
    expect(text.paragraphs.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Notes", "for June"])
  })

  it("keeps charts in the same stacking list", () => {
    const { c, o } = setup()
    o.insertShape("rect")
    c.charts.insert("ColumnClustered")
    o.insertShape("ellipse")
    expect(c.sheet.objects.map((x) => (isChart(x) ? "chart" : x.kind))).toEqual(["shape", "chart", "shape"])
    expect(c.sheet.charts).toHaveLength(1)
    expect(new Set(ids(c)).size).toBe(3)
  })
})

describe("editing objects", () => {
  it("moves, nudges, resizes, rotates and flips the selection", () => {
    const { c, o } = setup()
    const s = o.insertShape("rect", { x: 100, y: 100, w: 50, h: 40 })
    o.nudge(8, -8)
    expect(chartBox(c.sheet, o.find(s.id)!.anchor)).toMatchObject({ x: 108, y: 92 })
    o.resize(100, null, true)
    expect(chartBox(c.sheet, o.find(s.id)!.anchor)).toMatchObject({ w: 100, h: 80 })
    o.rotate(90)
    o.rotate(300)
    expect((o.find(s.id) as SheetShape).rotation).toBe(30)
    o.rotate(0, true)
    expect((o.find(s.id) as SheetShape).rotation ?? 0).toBe(0)
    o.flip("h")
    expect((o.find(s.id) as SheetShape).flipH).toBe(true)
  })

  it("changes the stacking order", () => {
    const { c, o } = setup()
    const a = o.insertShape("rect")
    const b = o.insertShape("ellipse")
    const d = o.insertShape("triangle")
    o.select(a.id)
    o.order("forward")
    expect(ids(c)).toEqual([b.id, a.id, d.id])
    o.order("front")
    expect(ids(c)).toEqual([b.id, d.id, a.id])
    o.select(d.id)
    o.order("back")
    expect(ids(c)).toEqual([d.id, b.id, a.id])
    o.order("backward")
    expect(ids(c)).toEqual([d.id, b.id, a.id])
    o.moveTo(d.id, 2)
    expect(ids(c)).toEqual([b.id, a.id, d.id])
  })

  it("aligns and distributes several objects", () => {
    const { c, o } = setup()
    const a = o.insertShape("rect", { x: 10, y: 10, w: 20, h: 20 })
    const b = o.insertShape("rect", { x: 100, y: 40, w: 40, h: 20 })
    const d = o.insertShape("rect", { x: 50, y: 90, w: 20, h: 20 })
    o.select([a.id, b.id, d.id])
    o.align("left")
    expect([a, b, d].map((x) => chartBox(c.sheet, o.find(x.id)!.anchor).x)).toEqual([10, 10, 10])
    o.align("bottom")
    expect([a, b, d].map((x) => Math.round(chartBox(c.sheet, o.find(x.id)!.anchor).y + 20))).toEqual([110, 110, 110])
    o.align("distributeH")
    const xs = [a, b, d].map((x) => chartBox(c.sheet, o.find(x.id)!.anchor)).sort((p, q) => p.x - q.x)
    expect(xs[0].x).toBe(10)
  })

  it("groups and ungroups, keeping each shape's place", () => {
    const { c, o } = setup()
    const a = o.insertShape("rect", { x: 10, y: 10, w: 20, h: 20 })
    const b = o.insertShape("ellipse", { x: 60, y: 40, w: 40, h: 30 })
    o.select([a.id, b.id])
    o.group()
    expect(c.sheet.objects).toHaveLength(1)
    const g = c.sheet.objects[0]
    expect(!isChart(g) && g.kind).toBe("group")
    expect(chartBox(c.sheet, g.anchor)).toMatchObject({ x: 10, y: 10, w: 90, h: 60 })
    o.ungroup()
    expect(c.sheet.objects.map((x) => x.name)).toEqual(["Rectangle 1", "Oval 2"])
    expect(chartBox(c.sheet, o.find(b.id)!.anchor)).toMatchObject({ x: 60, y: 40, w: 40, h: 30 })
    c.undo()
    expect(c.sheet.objects).toHaveLength(1)
  })

  it("duplicates, deletes, copies and pastes", () => {
    const { c, o } = setup()
    o.insertShape("rect")
    o.duplicate()
    expect(names(c)).toEqual(["Rectangle 1", "Rectangle 2"])
    const clip = o.copy(false, () => null)!
    expect(clip.html).toContain("cells-objects:")
    c.goTo({ r0: 10, c0: 3, r1: 10, c1: 3 })
    expect(o.paste(clip.html)).toBe(true)
    expect(names(c)).toEqual(["Rectangle 1", "Rectangle 2", "Rectangle 3"])
    expect(chartBox(c.sheet, c.sheet.objects[2].anchor).x).toBe(c.sheet.cols.offset(3))
    expect(o.paste("<p>other</p>")).toBe(false)
    o.delete()
    expect(names(c)).toHaveLength(2)
  })

  it("handles keys only while objects are selected", () => {
    const { c, o } = setup()
    expect(o.key("Delete", false, false)).toBe(false)
    o.insertShape("rect", { x: 40, y: 40, w: 20, h: 20 })
    expect(o.key("ArrowRight", false, false)).toBe(true)
    expect(chartBox(c.sheet, c.sheet.objects[0].anchor).x).toBe(48)
    expect(o.key("Escape", false, false)).toBe(true)
    expect(o.selected).toEqual([])
    o.select(c.sheet.objects[0].id)
    expect(o.key("Delete", false, false)).toBe(true)
    expect(c.sheet.objects).toHaveLength(0)
  })

  it("hides, renames and describes objects from the Selection Pane", () => {
    const { c, o } = setup()
    const s = o.insertShape("rect")
    o.setHidden(s.id, true)
    expect((o.find(s.id) as SheetShape).hidden).toBe(true)
    o.setAllHidden(false)
    expect((o.find(s.id) as SheetShape).hidden).toBeFalsy()
    o.rename(s.id, "Logo box")
    o.setAltText(s.id, "  Company logo ", false)
    expect(o.find(s.id)).toMatchObject({ name: "Logo box", altText: "Company logo" })
    o.setAltText(s.id, "", true)
    expect(o.find(s.id)).toMatchObject({ decorative: true })
    expect(o.find(s.id)!.altText).toBeUndefined()
    c.undo()
    expect(o.find(s.id)!.altText).toBe("Company logo")
  })
})

describe("objects elsewhere in Cells", () => {
  it("copies a sheet's objects with ids of their own", () => {
    const { c, o } = setup()
    o.insertShape("rect")
    o.insertPicture(PNG, { w: 10, h: 10 })
    c.sheets.copy(0, 1)
    const copy = c.wb.sheets[1].objects
    expect(copy.map((x) => x.name)).toEqual(["Rectangle 1", "Picture 2"])
    expect(copy.map((x) => x.id)).toEqual([3, 4])
  })

  it("counts pictures in Workbook Statistics and asks for their Alt Text", async () => {
    const { workbookStatistics } = await import("./statistics")
    const { checkAccessibility } = await import("./accessibility")
    const { c, o } = setup()
    const pic = o.insertPicture(PNG, { w: 10, h: 10 })
    o.insertShape("rect")
    const labelled = o.insertShape("rect")
    o.setText(labelled.id, "Total")
    expect(workbookStatistics(c.wb, 0).sheet.images).toBe(1)
    const alt = checkAccessibility(c.wb).filter((i) => i.kind === "altText")
    expect(alt.map((i) => i.label)).toEqual(["Picture 1", "Rectangle 2"])
    o.setAltText(pic.id, "Company logo", false)
    expect(checkAccessibility(c.wb).filter((i) => i.kind === "altText")).toHaveLength(1)
  })
})

describe("Picture Format › Crop", () => {
  it("crops in a mode Esc ends, each change one undo step, also on a flipped picture", () => {
    const { c, o } = setup()
    const pic = o.insertPicture(PNG, { w: 200, h: 100 })
    o.toggleCrop()
    expect(o.cropping).toBe(pic.id)
    const box = o.box(pic)
    o.setCrop(pic.id, { l: 0.25, t: 0, r: 0, b: 0 }, { ...box, x: box.x + 50, w: 150 })
    const cropped = o.picture()!
    expect(cropped.crop).toEqual({ l: 0.25, t: 0, r: 0, b: 0 })
    expect(chartBox(c.sheet, cropped.anchor).w).toBe(150)
    expect(o.key("Escape", false, false)).toBe(true)
    expect(o.cropping).toBeNull()
    expect(o.selected).toEqual([pic.id])
    c.undo()
    expect(o.picture()!.crop).toBeUndefined()
    // A flipped picture keeps its crop in its own sides: the screen's left is its right.
    o.flip("h")
    o.setCrop(pic.id, { l: 0.25, t: 0, r: 0, b: 0 }, box)
    expect(o.picture()!.crop).toEqual({ l: 0, t: 0, r: 0.25, b: 0 })
    expect(o.seenCrop(o.picture()!)).toEqual({ l: 0.25, t: 0, r: 0, b: 0 })
  })

  it("fills, fits, crops to an aspect ratio and to a shape; selecting away ends cropping", () => {
    const { c, o } = setup()
    const pic = o.insertPicture(PNG, { w: 100, h: 100 })
    o.resize(200, 100)
    o.fillOrFit("fill")
    expect(o.picture()!.crop).toEqual({ l: 0, t: 0.25, r: 0, b: 0.25 })
    o.fillOrFit("fit")
    expect(o.picture()!.crop).toEqual({ l: -0.5, t: 0, r: -0.5, b: 0 })
    c.undo()
    c.undo()
    o.cropToRatio(1, 1)
    expect(chartBox(c.sheet, o.picture()!.anchor)).toMatchObject({ w: 100, h: 100 })
    expect(o.picture()!.crop).toEqual({ l: 0.25, t: 0, r: 0.25, b: 0 })
    expect(o.cropping).toBe(pic.id)
    o.cropToShape("ellipse")
    expect(o.picture()!.geometry).toBe("ellipse")
    o.cropToShape("rect")
    expect(o.picture()!.geometry).toBeUndefined()
    o.select(null)
    expect(o.cropping).toBeNull()
  })

  it("can't crop a 3D model, a signature line or an embedded object's icon", () => {
    const { o } = setup()
    o.insertPicture(PNG, { w: 10, h: 10 }, { signature: { signer: "A" } as never })
    expect(o.croppable()).toBeNull()
    o.toggleCrop()
    expect(o.cropping).toBeNull()
  })
})

describe("Picture Format › Corrections, Color and Transparency", () => {
  const stand = (src: string, a: object) => Promise.resolve(`${src}#${JSON.stringify(a)}`)

  it("bakes corrections into the picture from its original, as one undo step each", async () => {
    const { c, o } = setup()
    o.correct = stand
    o.insertPicture(PNG, { w: 40, h: 20 })
    expect(await o.setAdjust({ bright: 0.2, contrast: -0.2 }, "Brightness/Contrast")).toBe(true)
    const p = o.picture()!
    expect(p.adjust).toEqual({ bright: 0.2, contrast: -0.2 })
    expect(p.original).toBe(PNG)
    expect(p.src).toBe(`${PNG}#{"bright":0.2,"contrast":-0.2}`)
    // A second correction starts again from the original, not the corrected picture.
    await o.setAdjust({ saturation: 2 }, "Color Saturation")
    expect(o.picture()!.src).toBe(`${PNG}#{"bright":0.2,"contrast":-0.2,"saturation":2}`)
    // Colour and transparency are drawn, not baked in: the picture stays.
    const before = o.picture()!.src
    await o.setAdjust({ transparency: 0.5, recolor: { kind: "grayscale" } }, "Transparency")
    expect(o.picture()!.src).toBe(before)
    expect(o.picture()!.adjust).toMatchObject({ transparency: 0.5, recolor: { kind: "grayscale" } })
    // Back to neutral: the original is the picture again.
    await o.setAdjust({ bright: 0, contrast: 0, saturation: 1, transparency: 0, recolor: undefined }, "Reset")
    expect(o.picture()!.src).toBe(PNG)
    expect(o.picture()!.original).toBeUndefined()
    expect(o.picture()!.adjust).toBeUndefined()
    c.undo()
    expect(o.picture()!.adjust).toMatchObject({ transparency: 0.5 })
  })

  it("lets a later change win while a correction is being made", async () => {
    const { o } = setup()
    let release: () => void = () => {}
    o.correct = (src, a) =>
      a.bright === 0.4 ? new Promise((done) => (release = () => done(stand(src, a)))) : stand(src, a)
    o.insertPicture(PNG, { w: 40, h: 20 })
    const slow = o.setAdjust({ bright: 0.4 }, "Brightness")
    const fast = o.setAdjust({ bright: 0.2 }, "Brightness")
    expect(await fast).toBe(true)
    release()
    expect(await slow).toBe(false)
    expect(o.picture()!.adjust).toEqual({ bright: 0.2 })
  })

  it("keeps Excel's corrections when their original can't be opened, but still changes transparency", async () => {
    const { o } = setup()
    o.correct = stand
    o.insertPicture(PNG, { w: 40, h: 20 })
    o.update("Excel", (x) => {
      if (isChart(x) || x.kind !== "picture") return
      x.adjust = { bright: 0.2 }
      x.original = "data:image/vnd.ms-photo;base64,AA=="
    })
    expect(o.canCorrect()).toBe(false)
    expect(await o.setAdjust({ bright: 0.4 }, "Brightness")).toBe(false)
    expect(await o.setAdjust({ transparency: 0.3 }, "Transparency")).toBe(true)
    expect(o.picture()!.adjust).toEqual({ bright: 0.2, transparency: 0.3 })
  })

  it("resets the picture to how it came, with or without its size", async () => {
    const { o } = setup()
    o.correct = stand
    o.insertPicture(PNG, { w: 40, h: 20 })
    await o.setAdjust({ sharpen: 0.25, transparency: 0.5 }, "Sharpen")
    o.cropToShape("ellipse")
    o.resetPicture(false)
    const p = o.picture()!
    expect([p.src, p.adjust, p.original, p.geometry]).toEqual([PNG, undefined, undefined, undefined])
  })

  it("only adjusts pictures, not icons", () => {
    const { o } = setup()
    o.insertPicture("data:image/svg+xml;base64,PHN2Zy8+", { w: 10, h: 10 })
    expect(o.adjustable()).toBeNull()
  })
})

describe("Picture Format › Artistic Effects", () => {
  it("bakes an effect in at the picture's size on the sheet at 220 dpi, as Excel works on it", async () => {
    const { o } = setup()
    const sizes: (number | undefined)[] = []
    o.correct = (src, a, side) => (sizes.push(side), Promise.resolve(`${src}#${JSON.stringify(a)}`))
    o.insertPicture(PNG, { w: 192, h: 96 })
    await o.setAdjust({ artistic: { effect: "Blur" } }, "Artistic Effects")
    expect(o.picture()!.adjust).toEqual({ artistic: { effect: "Blur" } })
    expect(o.picture()!.original).toBe(PNG)
    await o.setAdjust({ bright: 0.2 }, "Brightness")
    expect(sizes).toEqual([440, 440])
    await o.setAdjust({ artistic: undefined }, "Artistic Effects")
    expect(sizes[2]).toBeUndefined()
    expect(o.picture()!.adjust).toEqual({ bright: 0.2 })
  })
})
