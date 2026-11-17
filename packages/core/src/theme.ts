import type { ShapeColor, ThemeSlot } from "./shapes"

/**
 * The workbook's theme (Page Layout › Themes): its colours, which theme colours in shapes and charts follow, and
 * its fonts (headings and body). Cells' default is Office's 2023+ theme, which Excel 365 gives new workbooks.
 */
export type Theme = {
  name: string
  colors: Record<ThemeSlot, string>
  fonts: { major: string; minor: string }
}

export const OFFICE_THEME: Theme = {
  name: "Office Theme",
  colors: {
    dk1: "000000",
    lt1: "FFFFFF",
    dk2: "0E2841",
    lt2: "E8E8E8",
    accent1: "156082",
    accent2: "E97132",
    accent3: "196B24",
    accent4: "0F9ED5",
    accent5: "A02B93",
    accent6: "4EA72E",
    hlink: "467886",
    folHlink: "96607D",
  },
  fonts: { major: "Aptos Display", minor: "Aptos Narrow" },
}

export const copyTheme = (t: Theme): Theme => ({ name: t.name, colors: { ...t.colors }, fonts: { ...t.fonts } })

/** A colour as "#rrggbb" with the theme's colours, and its opacity (1: solid). */
export function colorCss(c: ShapeColor, theme: Theme): { hex: string; alpha: number } {
  let hex = (c.theme ? theme.colors[c.theme] : c.rgb) ?? "000000"
  if (c.lumMod !== undefined || c.lumOff !== undefined) hex = luminance(hex, c.lumMod ?? 1, c.lumOff ?? 0)
  return { hex: `#${hex.toUpperCase()}`, alpha: c.alpha ?? 1 }
}

/** DrawingML's lumMod/lumOff: the colour's lightness (HSL) times mod plus off. */
function luminance(hex: string, mod: number, off: number): string {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  let l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = Math.min(1, Math.max(0, l * mod + off))
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  let out: number[]
  if (s === 0) out = [l, l, l]
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    out = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)]
  }
  return out
    .map((v) =>
      Math.round(v * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
}
