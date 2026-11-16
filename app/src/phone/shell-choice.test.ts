import { describe, expect, it } from "vitest"
import { heightWithoutKeyboard, shellFor } from "./shell-choice"

describe("choosing the shell", () => {
  it.each([
    // width, height, finger, shell
    [360, 780, true, "phone"],
    [780, 360, true, "phone"], // a phone in landscape keeps the phone shell
    [599, 900, true, "phone"],
    [600, 960, true, "full"], // a small tablet gets the ribbon
    [1280, 800, true, "full"],
    [400, 700, false, "full"], // a narrow desktop window with a mouse
    [1920, 1080, false, "full"],
    [640, 800, true, "full"], // half of a tablet's split screen
    [500, 800, true, "phone"], // a narrower split
  ] as const)("%i × %i, finger %s → %s", (width, height, coarse, shell) => {
    expect(shellFor({ width, height, coarse })).toBe(shell)
  })
})

describe("the on-screen keyboard", () => {
  it("doesn't count: a tablet in landscape keeps the ribbon while typing", () => {
    // 1280 × 800 with a 455 px keyboard leaves the page 345 px tall.
    const height = heightWithoutKeyboard(345, 455)
    expect(height).toBe(800)
    expect(shellFor({ width: 1280, height, coarse: true })).toBe("full")
    expect(heightWithoutKeyboard(700, undefined)).toBe(700)
    expect(heightWithoutKeyboard(700, -3)).toBe(700)
  })
})
