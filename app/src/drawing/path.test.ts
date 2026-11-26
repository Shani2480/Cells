import { describe, expect, it } from "vitest"
import type { ShapePath, SheetShape } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { deletePoint, fitPaths, movePoint, parseSvgPath, pathSvg, pathVertices, presetAsPaths } from "./path"

const square: ShapePath = {
  w: 100,
  h: 100,
  commands: [
    { t: "M", x: 0, y: 0 },
    { t: "L", x: 100, y: 0 },
    { t: "L", x: 100, y: 100 },
    { t: "L", x: 0, y: 100 },
    { t: "Z" },
  ],
}

describe("shapes' own outlines", () => {
  it("stretch over the box they're drawn in", () => {
    expect(pathSvg(square, 200, 50)).toBe("M0 0 L200 0 L200 50 L0 50 Z")
  })

  it("read presets' SVG, arcs and quadratic curves as cubics", () => {
    const c = parseSvgPath("M0 50 A50 50 0 0 1 100 50 Q100 100 50 100 H0 V50 Z")
    expect(c[0]).toEqual({ t: "M", x: 0, y: 50 })
    const arcEnd = c.filter((x) => x.t === "C")[1] as { x: number; y: number }
    expect(arcEnd.x).toBeCloseTo(100)
    expect(arcEnd.y).toBeCloseTo(50)
    expect(c.slice(-3)).toEqual([{ t: "L", x: 0, y: 100 }, { t: "L", x: 0, y: 50 }, { t: "Z" }])
    // An oval in Edit Points: its four quarter curves, ending where they began.
    const [oval] = presetAsPaths("ellipse", 100, 60)
    expect(oval.commands.filter((x) => x.t === "C").length).toBeGreaterThanOrEqual(4)
  })

  it("show each point once (a closed figure's start isn't shown twice)", () => {
    const closed: ShapePath = {
      ...square,
      commands: [...square.commands.slice(0, 4), { t: "L", x: 0, y: 0 }, { t: "Z" }],
    }
    expect(pathVertices([closed], 100, 100)).toHaveLength(4)
  })

  it("move a point with its curve handles, and the box fits the points again", () => {
    const moved = movePoint([square], { path: 0, command: 2 }, 150, 120, 100, 100)
    expect(moved[0].commands[2]).toEqual({ t: "L", x: 150, y: 120 })
    const fit = fitPaths(moved, 100, 100)
    expect(fit.box).toEqual({ x: 0, y: 0, w: 150, h: 120 })
    const curve: ShapePath = {
      w: 10,
      h: 10,
      commands: [
        { t: "M", x: 0, y: 0 },
        { t: "C", x1: 2, y1: 0, x2: 8, y2: 10, x: 10, y: 10 },
        { t: "C", x1: 10, y1: 12, x2: 0, y2: 12, x: 0, y: 0 },
      ],
    }
    const c = movePoint([curve], { path: 0, command: 1 }, 12, 10, 10, 10)[0].commands
    expect(c[1]).toMatchObject({ x2: 10, x: 12 })
    expect(c[2]).toMatchObject({ x1: 12 })
  })

  it("delete a point, a figure keeping two", () => {
    const less = deletePoint([square], { path: 0, command: 2 })
    expect(less[0].commands.map((c) => c.t)).toEqual(["M", "L", "L", "Z"])
    const start = deletePoint([square], { path: 0, command: 0 })
    expect(start[0].commands[0]).toEqual({ t: "M", x: 100, y: 0 })
    const two: ShapePath = {
      w: 1,
      h: 1,
      commands: [
        { t: "M", x: 0, y: 0 },
        { t: "L", x: 1, y: 1 },
      ],
    }
    expect(deletePoint([two], { path: 0, command: 1 })[0]).toBe(two)
  })
})

describe("freeforms and Edit Points on the sheet", () => {
  it("a freeform drawn back to its start is a filled shape; a scribble is a line of curves", () => {
    const c = new SheetController()
    const pts = [
      { x: 100, y: 100 },
      { x: 300, y: 120 },
      { x: 200, y: 260 },
    ]
    const shape = c.objects.drawFreeform(pts, true, false)!
    expect(shape.geometry).toBe("custom")
    expect(c.objects.box(shape)).toMatchObject({ x: 100, y: 100, w: 200, h: 160 })
    expect(shape.fill.kind).toBe("solid")
    expect(shape.custom![0].commands.map((x) => x.t)).toEqual(["M", "L", "L", "Z"])
    const line = c.objects.drawFreeform(
      Array.from({ length: 30 }, (_, i) => ({ x: 400 + i * 5, y: 100 + Math.sin(i / 3) * 20 })),
      false,
      true,
    )!
    expect(line.fill.kind).toBe("none")
    expect(line.custom![0].fill).toBe(false)
    expect(line.custom![0].commands.slice(1).every((x) => x.t === "C")).toBe(true)
    c.undo()
    c.undo()
    expect(c.sheet.objects).toHaveLength(0)
  })

  it("Edit Points turns a preset into its points once one moves; the box follows; Escape ends it", () => {
    const c = new SheetController()
    const s = c.objects.insertShape("rect", { x: 96, y: 96, w: 100, h: 50 })
    c.objects.editPoints()
    expect(c.objects.editingPoints).toBe(s.id)
    const paths = c.objects.pointsOf(s)
    const corner = pathVertices(paths, 100, 50).find((v) => v.x === 100 && v.y === 50)!
    c.objects.setPoints(s.id, movePoint(paths, corner, 140, 80, 100, 50))
    const now = c.objects.find(s.id) as SheetShape
    expect(now.geometry).toBe("custom")
    expect(c.objects.box(now)).toMatchObject({ x: 96, y: 96, w: 140, h: 80 })
    expect(c.objects.key("Escape", false, false)).toBe(true)
    expect(c.objects.editingPoints).toBeNull()
    c.undo()
    expect((c.objects.find(s.id) as SheetShape).geometry).toBe("rect")
  })
})
