import { describe, expect, it } from "vitest"
import { isChart, type SheetShape } from "@cells/core"
import { SheetController } from "./controller"
import { chartBox } from "./charts"

const inks = (c: SheetController) =>
  c.sheet.objects.filter((o): o is SheetShape => !isChart(o) && o.kind === "shape" && o.ink !== undefined)

describe("ink (Draw tab)", () => {
  it("a pen stroke becomes an ink object spanning its points; strokes soon after join it", () => {
    const c = new SheetController()
    const o = c.objects
    o.usePen(1)
    expect(o.pen).toMatchObject({ kind: "pen", color: "#E71224" })
    o.addStroke([100, 100, 200, 150], 1000)
    expect(inks(c)).toHaveLength(1)
    expect(inks(c)[0].name).toBe("Ink 1")
    expect(chartBox(c.sheet, inks(c)[0].anchor)).toMatchObject({ x: 100, y: 100, w: 100, h: 50 })
    // A second stroke a second later joins the first; one much later makes a new ink object.
    o.addStroke([150, 300, 160, 310], 2000)
    expect(inks(c)).toHaveLength(1)
    expect(inks(c)[0].ink!.strokes).toHaveLength(2)
    expect(chartBox(c.sheet, inks(c)[0].anchor)).toMatchObject({ x: 100, y: 100, w: 100, h: 210 })
    o.addStroke([400, 400, 420, 420], 9000)
    expect(inks(c)).toHaveLength(2)
    // Each stroke is one undo step.
    c.undo()
    expect(inks(c)).toHaveLength(1)
    c.undo()
    expect(inks(c)[0].ink!.strokes).toHaveLength(1)
  })

  it("the highlighter's strokes are marked; a pen's colour and thickness can change", () => {
    const c = new SheetController()
    const o = c.objects
    o.setPenLook(0, { color: "#0070C0", width: 4 })
    expect(o.penSlot).toBe(0)
    o.addStroke([10, 10, 50, 10], 0)
    o.usePen(2)
    o.addStroke([10, 30, 50, 30], 10_000)
    const [pen, marker] = inks(c).map((i) => i.ink!.strokes[0])
    expect(pen).toMatchObject({ color: "#0070C0", width: 4 })
    expect(pen.highlighter).toBeUndefined()
    expect(marker.highlighter).toBe(true)
  })

  it("the eraser removes the strokes it touches, and an emptied ink object goes", () => {
    const c = new SheetController()
    const o = c.objects
    o.usePen(0)
    o.addStroke([100, 100, 200, 100], 0)
    o.addStroke([100, 200, 200, 200], 500)
    o.usePen("eraser")
    const hits = o.inkHits([150, 98, 151, 99])
    expect([...hits.values()].map((s) => [...s])).toEqual([[0]])
    expect(o.eraseStrokes(hits)).toBe(true)
    expect(inks(c)[0].ink!.strokes).toHaveLength(1)
    expect(chartBox(c.sheet, inks(c)[0].anchor)).toMatchObject({ y: 200 })
    o.eraseStrokes(o.inkHits([150, 200]))
    expect(inks(c)).toHaveLength(0)
    c.undo()
    expect(inks(c)).toHaveLength(1)
  })

  it("Hide Ink hides it without deleting; Delete All Ink clears this sheet or every sheet", () => {
    const c = new SheetController()
    const o = c.objects
    o.usePen(0)
    o.addStroke([10, 10, 20, 20], 0)
    o.usePen(null)
    expect(o.pen).toBeNull()
    o.setInkHidden(true)
    expect(o.inkHidden).toBe(true)
    expect(o.hasInk()).toBe(true)
    expect(o.deleteAllInk(true)).toBe(true)
    expect(o.hasInk(true)).toBe(false)
    c.undo()
    expect(o.hasInk()).toBe(true)
  })
})
