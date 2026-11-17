import { OFFICE_THEME, SMARTART_LAYOUTS, type SmartArt } from "@cells/core"
import { describe, expect, it } from "vitest"
import { layoutSmartArt } from "./smartart"
import { estimate } from "./render"

const PX = 4 / 3
/** Excel's frame for these numbers: 300 × 200 points. */
const FRAME = { w: 300 * PX, h: 200 * PX }
const art = (key: string, n: number): SmartArt => ({
  layout: SMARTART_LAYOUTS.find((l) => l.key === key)!.id,
  nodes: Array.from({ length: n }, (_, i) => ({ text: `P${i}`, children: [] })),
})
/** Shapes as Excel reports them: left, top, width, height in points, and the turn. */
const pts = (key: string, n: number) =>
  layoutSmartArt(art(key, n), FRAME, OFFICE_THEME, estimate).map((s) => [
    s.geometry,
    ...[s.box!.x, s.box!.y, s.box!.w, s.box!.h].map((v) => Math.round((v / PX) * 10) / 10 || 0),
    s.rotation ?? 0,
  ])

describe("SmartArt layouts, as Excel lays out Cells' definitions", () => {
  it("Basic Block List: rows of blocks, a short last row centred", () => {
    expect(pts("list", 4)).toEqual([
      ["rect", 0, 7.1, 142.9, 85.7, 0],
      ["rect", 157.1, 7.1, 142.9, 85.7, 0],
      ["rect", 0, 107.1, 142.9, 85.7, 0],
      ["rect", 157.1, 107.1, 142.9, 85.7, 0],
    ])
    // Excel: 45.1,0.1 99.9x59.9 … 100,140 for the fifth.
    const five = pts("list", 5)
    expect(five[0]).toEqual(["rect", 45, 0, 100, 60, 0])
    expect(five[4]).toEqual(["rect", 100, 140, 100, 60, 0])
  })

  it("Basic Process: steps in a row, an arrow in each gap", () => {
    // Excel: 0,82.2 59.4x35.6; arrows 64.6,93.6 11x12.9.
    const four = pts("process", 4)
    expect(four[0]).toEqual(["roundRect", 0, 82.2, 59.4, 35.6, 0])
    expect(four[4]).toEqual(["rightArrow", 64.6, 93.6, 11, 12.9, 0])
    expect(four).toHaveLength(7)
  })

  it("Basic Cycle: circles on a ring from the top, arrows along it back to the first", () => {
    // Excel (3): 103.9,0 92.1; 166.2,107.8; 41.7,107.8; arrows 22.6x22.1 turned 60 and 180.
    const three = pts("cycle", 3)
    expect(three.slice(0, 3)).toEqual([
      ["ellipse", 103.9, 0, 92.2, 92.2, 0],
      // Excel: 166.2 (its rounding).
      ["ellipse", 166.1, 107.8, 92.2, 92.2, 0],
      ["ellipse", 41.7, 107.8, 92.2, 92.2, 0],
    ])
    expect(three[3]![3]).toBeCloseTo(22.6, 0)
    expect(three[3]![4]).toBeCloseTo(22.1, 0)
    expect(three.map((s) => s[5])).toEqual([0, 0, 0, 60, 180, 300])
  })

  it("Grid Matrix: squares two to a row", () => {
    expect(pts("matrix", 4)).toEqual([
      ["rect", 50, 0, 97.6, 97.6, 0],
      ["rect", 152.4, 0, 97.6, 97.6, 0],
      ["rect", 50, 102.4, 97.6, 97.6, 0],
      ["rect", 152.4, 102.4, 97.6, 97.6, 0],
    ])
    expect(pts("matrix", 3)[2]).toEqual(["rect", 101.2, 102.4, 97.6, 97.6, 0])
  })

  it("shrinks the text of every point together until the longest fits", () => {
    const shapes = layoutSmartArt(
      {
        ...art("list", 3),
        nodes: [
          { text: "A much longer piece of text", children: [] },
          { text: "B", children: [] },
          { text: "", children: [] },
        ],
      },
      FRAME,
      OFFICE_THEME,
      estimate,
    )
    const sizes = shapes.map((s) => s.text!.paragraphs[0]!.runs[0]!.size)
    expect(new Set(sizes).size).toBe(1)
    expect(sizes[0]).toBeLessThan(50)
    expect(shapes[2]!.text!.paragraphs[0]!.runs[0]!.text).toBe("[Text]")
  })
})
