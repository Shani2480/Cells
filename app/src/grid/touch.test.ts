import { describe, expect, it } from "vitest"
import { keyboardOverlap, pinch, TapTracker } from "./touch"

describe("touch gestures", () => {
  const clamp = (z: number) => Math.min(4, Math.max(0.1, z))

  it("pinch zooms by the change in finger distance, keeping the midpoint's cell in place", () => {
    const origin = { x: 40, y: 20 }
    const scroll = { x: 300, y: 500 }
    const mid = { x: 240, y: 220 }
    const out = pinch({ startZoom: 1, startDistance: 100, distance: 200, mid, scroll, origin, zoom: 1, clamp })
    expect(out.zoom).toBe(2)
    // The sheet point under the midpoint before (scroll + mid − origin) is under it after.
    const before = { x: scroll.x + mid.x - origin.x, y: scroll.y + mid.y - origin.y }
    expect(out.scroll.x + mid.x / 2 - origin.x).toBeCloseTo(before.x)
    expect(out.scroll.y + mid.y / 2 - origin.y).toBeCloseTo(before.y)
  })

  it("pinch stays within the zoom limits and never scrolls before A1", () => {
    const base = { startZoom: 1, startDistance: 100, mid: { x: 500, y: 400 }, origin: { x: 40, y: 20 }, zoom: 1, clamp }
    expect(pinch({ ...base, distance: 1000, scroll: { x: 0, y: 0 } }).zoom).toBe(4)
    const out = pinch({ ...base, distance: 1, scroll: { x: 0, y: 0 } })
    expect(out.zoom).toBe(0.1)
    expect(out.scroll).toEqual({ x: 0, y: 0 })
  })

  it("a double tap is two quick taps on one cell", () => {
    const taps = new TapTracker()
    expect(taps.tap(2, 3, 1000)).toBe(false)
    expect(taps.tap(2, 3, 1200)).toBe(true)
    // A third tap starts over; a slow or elsewhere tap doesn't count.
    expect(taps.tap(2, 3, 1300)).toBe(false)
    expect(taps.tap(2, 3, 2000)).toBe(false)
    expect(taps.tap(2, 4, 2100)).toBe(false)
  })

  it("scrolls the editor above the on-screen keyboard", () => {
    // The keyboard leaves 400 px of a 700 px window: an editor ending at 450 must move up 62 px (12 px margin).
    expect(keyboardOverlap(450, 0, 400)).toBe(62)
    expect(keyboardOverlap(300, 0, 400)).toBe(0)
  })
})
