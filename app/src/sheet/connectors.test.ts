import { describe, expect, it } from "vitest"
import type { SheetShape } from "@cells/core"
import { SheetController } from "./controller"
import { lineEnds, sitesOf } from "./connectors"

/**
 * Lines glued to shapes, as Excel's connectors: drawn from one shape's connection site to another's, they follow the
 * shapes as they move, turn or change size, in the same undo step; moving the line by itself lets go.
 */
function setup() {
  const c = new SheetController()
  const a = c.objects.insertShape("rect", { x: 100, y: 100, w: 100, h: 60 })
  const b = c.objects.insertShape("ellipse", { x: 400, y: 300, w: 80, h: 80 })
  c.objects.pending = { geometry: "straightConnector1" }
  // From near A's right site (200, 130) to near B's left site (400, 340).
  const line = c.objects.drawLine({ x: 204, y: 127 }, { x: 396, y: 343 })!
  const get = (id: number) => c.objects.find(id) as SheetShape
  const ends = () => lineEnds(get(line.id), c.objects.box(get(line.id))).map((p) => [Math.round(p.x), Math.round(p.y)])
  return { c, a, b, line, get, ends }
}

describe("connectors glued to shapes", () => {
  it("a line drawn onto shapes' sites snaps to them and is glued", () => {
    const { line, a, b, ends } = setup()
    expect(line.from).toEqual({ id: a.id, site: 3 })
    expect(line.to).toEqual({ id: b.id, site: 2 })
    expect(ends()).toEqual([
      [200, 130],
      [400, 340],
    ])
  })

  it("follows a glued shape that moves, turns or changes size, in the same undo step", () => {
    const { c, b, ends, get } = setup()
    c.objects.place(new Map([[b.id, { x: 50, y: 400, w: 80, h: 80 }]]))
    // B is now left of A's right side: the line runs back, flipped.
    expect(ends()).toEqual([
      [200, 130],
      [50, 440],
    ])
    c.undo()
    expect(ends()).toEqual([
      [200, 130],
      [400, 340],
    ])
    c.objects.select(b.id)
    c.objects.rotate(90)
    const left = sitesOf(get(b.id), c.objects.box(get(b.id)))[2]
    expect(ends()[1]).toEqual([Math.round(left.x), Math.round(left.y)])
    expect(ends()[1]).toEqual([440, 300])
  })

  it("moving the line by itself lets go; a shape deleted lets go of its end; an end dragged onto a site glues", () => {
    const { c, a, b, line, get } = setup()
    const box = c.objects.box(get(line.id))
    c.objects.place(new Map([[line.id, { ...box, x: box.x + 30 }]]))
    expect(get(line.id).from).toBeUndefined()
    expect(get(line.id).to).toBeUndefined()
    c.undo()
    c.objects.select(b.id)
    c.objects.delete()
    expect(get(line.id).to).toBeUndefined()
    expect(get(line.id).from).toEqual({ id: a.id, site: 3 })
    c.objects.moveLineEnd(line.id, "to", { x: 150, y: 162 })
    expect(get(line.id).to).toEqual({ id: a.id, site: 2 })
    c.objects.moveLineEnd(line.id, "to", { x: 600, y: 600 })
    expect(get(line.id).to).toBeUndefined()
  })
})
