import { describe, expect, it } from "vitest"
import { aspectCrop, cropFrame, cropOf, fillCrop, fitCrop, flipCrop, fullBox, isNoCrop, scaleFull } from "./crop"

const frame = { x: 100, y: 50, w: 200, h: 100 }

describe("Picture Format › Crop", () => {
  it("finds the whole picture from the frame and its crop, and the crop back from both", () => {
    const crop = { l: 0.25, t: 0, r: 0.25, b: 0.5 }
    const full = fullBox(frame, crop)
    expect(full).toEqual({ x: 0, y: 50, w: 400, h: 200 })
    expect(cropOf(frame, full)).toEqual(crop)
    expect(isNoCrop(cropOf(frame, fullBox(frame)))).toBe(true)
  })

  it("moves the frame's sides with the crop handles and keeps the picture where it is", () => {
    const full = fullBox(frame)
    const left = cropFrame(frame, "w", 50, 0)
    expect(left).toEqual({ x: 150, y: 50, w: 150, h: 100 })
    expect(cropOf(left, full)).toEqual({ l: 0.25, t: 0, r: 0, b: 0 })
    // Outward past the picture adds empty space (a negative crop), as Excel does.
    expect(cropOf(cropFrame(frame, "s", 0, 20), full).b).toBeCloseTo(-0.2)
    // Shift on a corner keeps the frame's proportions; a frame is never smaller than the least size.
    expect(cropFrame(frame, "se", -100, -10, true)).toEqual({ x: 100, y: 50, w: 100, h: 50 })
    expect(cropFrame(frame, "e", -500, 0).w).toBe(8)
  })

  it("sizes the whole picture from a corner in its proportions", () => {
    const full = { x: 0, y: 0, w: 400, h: 200 }
    expect(scaleFull(full, "se", 100, 0)).toEqual({ x: 0, y: 0, w: 500, h: 250 })
    expect(scaleFull(full, "se", 100, 100)).toEqual({ x: 0, y: 0, w: 600, h: 300 })
    expect(scaleFull(full, "nw", 200, 200)).toEqual({ x: 200, y: 100, w: 200, h: 100 })
    expect(scaleFull(full, "nw", 1000, 1000).h).toBe(8)
  })

  it("fills the frame or fits in it, centred, in the picture's own proportions", () => {
    // A square picture in a 2:1 frame: Fill cuts a quarter from top and bottom, Fit leaves a quarter each side empty.
    expect(fillCrop(frame, 1)).toEqual({ l: 0, t: 0.25, r: 0, b: 0.25 })
    expect(fitCrop(frame, 1)).toEqual({ l: -0.5, t: 0, r: -0.5, b: 0 })
  })

  it("crops to an aspect ratio from the picture's middle", () => {
    const { frame: square, crop } = aspectCrop(frame, { l: 0, t: 0, r: 0, b: 0 }, 1)
    expect(square).toEqual({ x: 150, y: 50, w: 100, h: 100 })
    expect(crop).toEqual({ l: 0.25, t: 0, r: 0.25, b: 0 })
  })

  it("turns a crop to the screen's sides for a flipped picture, and back", () => {
    const c = { l: 0.1, t: 0.2, r: 0.3, b: 0.4 }
    expect(flipCrop(c, true)).toEqual({ l: 0.3, t: 0.2, r: 0.1, b: 0.4 })
    expect(flipCrop(flipCrop(c, true, true), true, true)).toEqual(c)
  })
})
