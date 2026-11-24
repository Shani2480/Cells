import { describe, expect, it } from "vitest"
import type { PictureAdjust } from "@cells/core"
import { correctionsOf, correctPixels, loadCurves, sharpenSoften } from "./picture-adjust"

/** Pixels Excel 365 corrected (a setting, a colour before, and Excel's colour after), from its saved pictures. */
const EXCEL: [PictureAdjust, number[], number[]][] = [
  [{ bright: 0.2, contrast: 0.2 }, [200, 0, 0], [252, 0, 0]],
  [{ bright: 0.2, contrast: 0.2 }, [0, 64, 0], [0, 80, 0]],
  [{ bright: 0.2, contrast: 0.2 }, [150, 105, 194], [207, 147, 247]],
  [{ bright: 0.2, contrast: 0.2 }, [30, 225, 90], [32, 255, 124]],
  [{ bright: 0.2, contrast: 0.2 }, [0, 0, 240], [0, 0, 255]],
  [{ contrast: -0.4 }, [200, 0, 0], [172, 0, 0]],
  [{ contrast: -0.4 }, [0, 64, 0], [0, 89, 0]],
  [{ contrast: -0.4 }, [150, 105, 194], [140, 115, 168]],
  [{ contrast: -0.4 }, [30, 225, 90], [58, 197, 106]],
  [{ contrast: -0.4 }, [0, 0, 240], [0, 0, 220]],
  [{ bright: -0.4 }, [200, 0, 0], [111, 0, 0]],
  [{ bright: -0.4 }, [0, 64, 0], [0, 31, 0]],
  [{ bright: -0.4 }, [150, 105, 194], [75, 51, 107]],
  [{ bright: -0.4 }, [30, 225, 90], [15, 130, 44]],
  [{ bright: -0.4 }, [0, 0, 240], [0, 0, 141]],
  [{ saturation: 0.33 }, [200, 0, 0], [133, 67, 67]],
  [{ saturation: 0.33 }, [0, 64, 0], [21, 43, 21]],
  [{ saturation: 0.33 }, [150, 105, 194], [150, 135, 164]],
  [{ saturation: 0.33 }, [30, 225, 90], [95, 160, 115]],
  [{ saturation: 0.33 }, [0, 0, 240], [80, 80, 160]],
  [{ saturation: 2 }, [200, 0, 0], [200, 0, 0]],
  [{ saturation: 2 }, [0, 64, 0], [0, 64, 0]],
  [{ saturation: 2 }, [150, 105, 194], [150, 81, 218]],
  [{ saturation: 2 }, [30, 225, 90], [16, 239, 85]],
  [{ saturation: 2 }, [0, 0, 240], [0, 0, 240]],
  [{ temperature: 4700 }, [200, 0, 0], [182, 0, 0]],
  [{ temperature: 4700 }, [0, 64, 0], [0, 65, 0]],
  [{ temperature: 4700 }, [150, 105, 194], [135, 107, 226]],
  [{ temperature: 4700 }, [30, 225, 90], [26, 228, 112]],
  [{ temperature: 4700 }, [0, 0, 240], [0, 0, 255]],
  [{ temperature: 11200 }, [200, 0, 0], [221, 0, 0]],
  [{ temperature: 11200 }, [0, 64, 0], [0, 63, 0]],
  [{ temperature: 11200 }, [150, 105, 194], [168, 104, 163]],
  [{ temperature: 11200 }, [30, 225, 90], [35, 223, 71]],
  [{ temperature: 11200 }, [0, 0, 240], [0, 0, 205]],
]

describe("Picture Format › Corrections and Color, as Excel does them", () => {
  it("gives Excel's colours for brightness, contrast, saturation and colour tone", async () => {
    const curves = await loadCurves()
    for (const [adjust, before, after] of EXCEL) {
      const data = new Uint8ClampedArray([...before, 255])
      correctPixels(data, 1, 1, adjust, curves)
      // Excel's pictures vary by a level or two from one to the next (the curves were read from others).
      for (let ch = 0; ch < 3; ch++)
        expect(Math.abs(data[ch] - after[ch]), JSON.stringify(adjust)).toBeLessThanOrEqual(3)
    }
  })

  it("sharpens edges and softens them, and leaves a flat picture alone", () => {
    // A 4×1 picture: dark, dark, light, light.
    const edge = () => new Uint8ClampedArray([80, 80, 80, 255, 80, 80, 80, 255, 176, 176, 176, 255, 176, 176, 176, 255])
    const sharp = edge()
    sharpenSoften(sharp, 4, 1, 1)
    expect([sharp[4], sharp[8]]).toEqual([18, 238])
    const soft = edge()
    sharpenSoften(soft, 4, 1, -0.5)
    expect(soft[4]).toBeGreaterThan(80)
    expect(soft[8]).toBeLessThan(176)
    const flat = new Uint8ClampedArray(16).fill(120)
    sharpenSoften(flat, 2, 2, 1)
    expect([...flat]).toEqual(new Array(16).fill(120))
  })

  it("keeps only the corrections that change something", () => {
    expect(correctionsOf({ saturation: 1, temperature: 6500, transparency: 0.5 })).toBeNull()
    expect(correctionsOf({ bright: 0.2, saturation: 1, recolor: { kind: "grayscale" } })).toEqual({ bright: 0.2 })
  })
})

describe("Picture Format › Color and Transparency, as drawn", () => {
  it("draws recolouring, washout and a transparent colour as a filter, and transparency as opacity", async () => {
    const { pictureFilter, pictureOpacity, duotoneRgb } = await import("./picture-filter")
    const { OFFICE_THEME } = await import("@cells/core")
    expect(pictureFilter("f", undefined, OFFICE_THEME)).toBe("")
    expect(pictureFilter("f", { transparency: 0.5 }, OFFICE_THEME)).toBe("")
    expect(pictureOpacity({ transparency: 0.3 })).toBeCloseTo(0.7)
    const grey = pictureFilter("f", { recolor: { kind: "grayscale" } }, OFFICE_THEME)
    expect(grey).toContain('<filter id="f"')
    expect(grey).toContain("0.2126 0.7152 0.0722")
    // Washout as Excel draws it: black becomes 205, and each level adds 0.3.
    const washout = pictureFilter("f", { lum: { bright: 0.7, contrast: -0.7 } }, OFFICE_THEME)
    expect(washout).toContain('slope="0.3" intercept="0.805"')
    // Brightness +30%, contrast +40%: Excel's black is 16, its slope 1.67.
    expect(pictureFilter("f", { lum: { bright: 0.3, contrast: 0.4 } }, OFFICE_THEME)).toContain(
      'slope="1.6667" intercept="0.0667"',
    )
    const key = pictureFilter("f", { transparentColor: "FF8000", recolor: { kind: "grayscale" } }, OFFICE_THEME)
    expect(key.indexOf("discrete")).toBeLessThan(key.indexOf("0.2126"))
    expect(key).toContain('in="s0"')
    // Sepia's light colour: D9C3A5 tinted half way to white in linear light, its saturation raised.
    const [r, g, b] = duotoneRgb({ rgb: "D9C3A5", tint: 0.5, satMod: 1.8 }, OFFICE_THEME).map((v) =>
      Math.round(v * 255),
    )
    expect(r).toBeGreaterThan(g)
    expect(g).toBeGreaterThan(b)
    expect(duotoneRgb({ preset: "white" }, OFFICE_THEME)).toEqual([1, 1, 1])
  })
})

describe("Picture Format › Artistic Effects", () => {
  // A small picture with an edge, a gradient and colour.
  const picture = () => {
    const w = 24
    const h = 16
    const data = new Uint8ClampedArray(w * h * 4)
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4
        data.set(x < 12 ? [200, 60 + y * 8, 40, 255] : [40, 120, 220 - x * 4, 255], i)
      }
    return { data, w, h }
  }

  it("changes the picture with each of Office's 22 effects, the same way each time, keeping its transparency", async () => {
    const { ARTISTIC_EFFECTS } = await import("@cells/core")
    const { applyArtistic } = await import("./picture-artistic")
    const effects = Object.keys(ARTISTIC_EFFECTS) as (keyof typeof ARTISTIC_EFFECTS)[]
    expect(effects).toHaveLength(22)
    for (const effect of effects) {
      const a = picture()
      const b = picture()
      applyArtistic(a.data, a.w, a.h, { effect })
      applyArtistic(b.data, b.w, b.h, { effect })
      expect([...a.data], effect).toEqual([...b.data])
      expect([...a.data], effect).not.toEqual([...picture().data])
      for (let i = 3; i < a.data.length; i += 4) expect(a.data[i]).toBe(255)
    }
  })

  it("lets the picture show through by its transparency", async () => {
    const { applyArtistic } = await import("./picture-artistic")
    const shown = picture()
    applyArtistic(shown.data, shown.w, shown.h, { effect: "Cutout", trans: 1 })
    expect([...shown.data]).toEqual([...picture().data])
    const half = picture()
    const full = picture()
    applyArtistic(half.data, half.w, half.h, { effect: "Cutout", trans: 0.5 })
    applyArtistic(full.data, full.w, full.h, { effect: "Cutout", trans: 0 })
    const original = picture().data
    expect(Math.abs(half.data[0] - (full.data[0] + original[0]) / 2)).toBeLessThanOrEqual(1)
  })
})
