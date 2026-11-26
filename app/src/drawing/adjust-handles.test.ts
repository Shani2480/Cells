import { describe, expect, it } from "vitest"
import { adjustHandles } from "./geometry"

const SHAPES = [
  "roundRect",
  "snip1Rect",
  "triangle",
  "parallelogram",
  "trapezoid",
  "hexagon",
  "octagon",
  "plus",
  "frame",
  "chevron",
  "homePlate",
  "donut",
  "can",
  "moon",
  "rightArrow",
  "leftArrow",
  "leftRightArrow",
  "upArrow",
  "downArrow",
  "wedgeRectCallout",
  "wedgeEllipseCallout",
]

describe("adjustment handles", () => {
  it("a handle dragged to a point stands there afterwards (every shape with handles)", () => {
    const w = 200
    const h = 120
    for (const name of SHAPES) {
      const handles = adjustHandles(name, w, h)
      expect(handles.length, name).toBeGreaterThan(0)
      handles.forEach((handle, i) => {
        // A point a little way from the handle towards the shape's middle (inside every handle's range).
        const tx = handle.x + (w / 2 - handle.x) * 0.15
        const ty = handle.y + (h / 2 - handle.y) * 0.15
        const adj = handle.drag(tx, ty)
        const moved = adjustHandles(name, w, h, adj)[i]
        // The handle follows the drag along the axis it moves on.
        const along = (a: number, b: number) => Math.abs(a - b) < 1
        expect(along(moved.x, tx) || along(moved.y, ty), `${name} #${i}`).toBe(true)
      })
    }
  })

  it("keeps values in Office's range and shapes without handles have none", () => {
    const [corner] = adjustHandles("roundRect", 200, 100)
    expect(corner.drag(500, 0)).toEqual({ adj: 50000 })
    expect(corner.drag(-20, 0)).toEqual({ adj: 0 })
    expect(adjustHandles("rect", 100, 100)).toEqual([])
    expect(adjustHandles("ellipse", 100, 100)).toEqual([])
    const [shaft, head] = adjustHandles("rightArrow", 200, 100)
    expect(shaft.drag(0, 25)).toEqual({ adj1: 50000 })
    expect(head.drag(150, 0)).toEqual({ adj2: 50000 })
  })
})
