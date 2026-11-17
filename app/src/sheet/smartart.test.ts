import { readFileSync } from "node:fs"
import { join } from "node:path"
import { addSmartArtNode, OFFICE_THEME, setSmartArtText, type SheetGroup, type SheetShape } from "@cells/core"
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

describe("Insert › SmartArt", () => {
  const fresh = () => {
    const c = new SheetController()
    const g = c.objects.insertSmartArt("urn:cells.app/smartart/layout/process")
    return { c, g, sa: () => c.objects.smartArt()! }
  }

  it("puts in a layout with its empty points, the Text Pane open, as one undo step", () => {
    const { c, g, sa } = fresh()
    expect(c.wb.undoLabel).toBe("Insert SmartArt")
    expect(sa().id).toBe(g.id)
    expect(sa().smartArt.nodes).toHaveLength(3)
    expect(c.objects.textPaneOpen).toBe(true)
    // Three steps, two arrows; "[Text]" where nothing's typed yet.
    expect(sa().items.map((i) => (i as SheetShape).geometry)).toEqual([
      "roundRect",
      "roundRect",
      "roundRect",
      "rightArrow",
      "rightArrow",
    ])
    expect((sa().items[0] as SheetShape).text!.paragraphs[0]!.runs[0]!.text).toBe("[Text]")
    expect(sa().smartArt.parts!.data).toContain('loTypeId="urn:cells.app/smartart/layout/process"')
    c.wb.undo()
    expect(c.sheet.objects).toHaveLength(0)
  })

  it("typing and Add Shape lay it out again; undo steps back", () => {
    const { c, g, sa } = fresh()
    c.objects.setSmartArt(g.id, { nodes: setSmartArtText(sa().smartArt.nodes, [0], "Plan") }, "Typing")
    expect((sa().items[0] as SheetShape).text!.paragraphs[0]!.runs[0]!.text).toBe("Plan")
    expect(sa().smartArt.parts!.data).toContain("<a:t>Plan</a:t>")
    c.objects.smartArtPoint = [2]
    expect(c.objects.editSmartArt((n, p) => addSmartArtNode(n, p), "Add Shape")).toBe(true)
    expect(sa().smartArt.nodes).toHaveLength(4)
    expect(c.objects.smartArtPoint).toEqual([3])
    expect(sa().items).toHaveLength(7)
    c.wb.undo()
    expect(sa().smartArt.nodes).toHaveLength(3)
  })

  it("Change Layout and Change Colors draw it anew; Convert to Shapes leaves a plain group", () => {
    const { c, g, sa } = fresh()
    c.objects.setSmartArt(g.id, { layout: "urn:cells.app/smartart/layout/cycle" }, "Change Layout")
    expect(sa().items.filter((i) => (i as SheetShape).geometry === "ellipse")).toHaveLength(3)
    c.objects.setSmartArt(g.id, { colors: "urn:cells.app/smartart/colors/accent2" }, "Change Colors")
    expect((sa().items[0] as SheetShape).fill).toEqual({ kind: "solid", color: { theme: "accent2" } })
    expect(sa().smartArt.parts!.colors).toContain('uniqueId="urn:cells.app/smartart/colors/accent2"')
    c.objects.convertToShapes(g.id)
    expect(c.objects.smartArt()).toBeNull()
    const plain = c.sheet.objects[0] as SheetGroup
    expect(plain.items).toHaveLength(6)
    expect(new Set(plain.items.map((i) => i.id)).size).toBe(6)
  })

  it("a SmartArt sized anew is laid out in its new frame", () => {
    const { c, g, sa } = fresh()
    const box = c.objects.box(g)
    c.objects.place(new Map([[g.id, { ...box, w: box.w * 2 }]]), "Size")
    expect(sa().child.w).toBeCloseTo(box.w * 2)
  })
})
