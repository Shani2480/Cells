import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { wheelRows } from "./wheel"

describe("the mouse wheel scrolls by rows (Excel)", () => {
  it("tells a wheel's notches from a touchpad", () => {
    expect(wheelRows({ deltaMode: 1, deltaX: 0, deltaY: 3 })).toBe(3)
    expect(wheelRows({ deltaMode: 0, deltaX: 0, deltaY: 100, wheelDeltaY: -120 })).toBe(3)
    expect(wheelRows({ deltaMode: 0, deltaX: 0, deltaY: -200, wheelDeltaY: 240 })).toBe(-6)
    expect(wheelRows({ deltaMode: 0, deltaX: 0, deltaY: 4.5, wheelDeltaY: -13 })).toBeNull()
    expect(wheelRows({ deltaMode: 0, deltaX: 2, deltaY: 100, wheelDeltaY: -120 })).toBeNull()
  })

  it("scrolls whole rows, skipping hidden ones, and lines up a row cut in half", () => {
    const c = new SheetController()
    c.scrollRows(3)
    expect(c.scrollY).toBe(60)
    c.sheet.rows.setHidden(4, true)
    // Rows 4 (hidden), 5, 6, 7: row 7 at the top, six rows down.
    c.scrollRows(3)
    expect(c.scrollY).toBe(6 * 20)
    c.scrollRows(-3)
    expect(c.scrollY).toBe(3 * 20)
    c.scrollTo(0, 50)
    c.scrollRows(3)
    expect(c.scrollY).toBe(4 * 20)
    c.scrollTo(0, 50)
    c.scrollRows(-3)
    expect(c.scrollY).toBe(0)
    c.scrollRows(-3)
    expect(c.scrollY).toBe(0)
  })
})
