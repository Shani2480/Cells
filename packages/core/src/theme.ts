import type { ShapeColor, ThemeSlot } from "./shapes"

/**
 * The workbook's theme (Page Layout › Themes): its colours, which theme colours in shapes and charts follow, and
 * its fonts (headings and body). Cells' default is Office's 2023+ theme, which Excel 365 gives new workbooks.
 */
export type Theme = {
  name: string
  colors: Record<ThemeSlot, string>
  fonts: { major: string; minor: string }
  /** The names of the colour and font sets (Page Layout › Colors and Fonts); the theme's name when absent. */
  colorsName?: string
  fontsName?: string
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

export const copyTheme = (t: Theme): Theme => ({
  ...t,
  colors: { ...t.colors },
  fonts: { ...t.fonts },
})

/** Whether two themes look the same (colours and fonts; names aside). */
export const sameLook = (a: Theme, b: Theme): boolean =>
  a.fonts.major === b.fonts.major &&
  a.fonts.minor === b.fonts.minor &&
  (Object.keys(a.colors) as ThemeSlot[]).every((k) => a.colors[k].toUpperCase() === b.colors[k].toUpperCase())

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

/**
 * Excel's tint: lighten (tint > 0) or darken (tint < 0) the colour's luminance. Excel works in Windows' integer HLS
 * (0–240, ColorRGBToHLS / ColorHLSToRGB) and truncates the new luminance; this matches the colours Excel 365 shows
 * (62 of 66 sampled theme tints exactly, the rest within 1 of 255).
 */
export function applyTint(hex: string, tint: number): string {
  const H = 240
  const M = 255
  const div = Math.trunc
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = div(((max + min) * H + M) / (2 * M))
  let h = (H * 2) / 3
  let s = 0
  if (max !== min) {
    const d = max - min
    s =
      l <= H / 2
        ? div((d * H + (max + min) / 2) / (max + min))
        : div((d * H + (2 * M - max - min) / 2) / (2 * M - max - min))
    const delta = (c: number) => div(((max - c) * (H / 6) + d / 2) / d)
    h = r === max ? delta(b) - delta(g) : g === max ? H / 3 + delta(r) - delta(b) : (2 * H) / 3 + delta(g) - delta(r)
    if (h < 0) h += H
    if (h > H) h -= H
  }
  const lum = div(tint < 0 ? l * (1 + tint) : l * (1 - tint) + (H - H * (1 - tint)))
  const to = (v: number) => Math.min(255, Math.max(0, v)).toString(16).padStart(2, "0")
  if (s === 0) {
    const v = to(div((lum * M) / H))
    return (v + v + v).toUpperCase()
  }
  const hue = (n1: number, n2: number, x: number) => {
    if (x < 0) x += H
    if (x > H) x -= H
    if (x < H / 6) return n1 + div(((n2 - n1) * x + H / 12) / (H / 6))
    if (x < H / 2) return n2
    if (x < (H * 2) / 3) return n1 + div(((n2 - n1) * ((H * 2) / 3 - x) + H / 12) / (H / 6))
    return n1
  }
  const m2 = lum <= H / 2 ? div((lum * (H + s) + H / 2) / H) : lum + s - div((lum * s + H / 2) / H)
  const m1 = 2 * lum - m2
  const channel = (x: number) => to(div((hue(m1, m2, x) * M + H / 2) / H))
  return (channel(h + H / 3) + channel(h) + channel(h - H / 3)).toUpperCase()
}

/** SpreadsheetML's theme colour order (`<color theme="n">`): lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink. */
export const SHEET_THEME_SLOTS: readonly ThemeSlot[] = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
]

/** A cell colour taken from the theme: its slot in SpreadsheetML's order and Excel's tint (-1 darker … 1 lighter). */
export type ThemeColorRef = { theme: number; tint?: number }

/** A theme colour of a cell as "#RRGGBB" under `theme`. */
export function themeColorHex(ref: ThemeColorRef, theme: Theme): string {
  const hex = theme.colors[SHEET_THEME_SLOTS[ref.theme] ?? "dk1"] ?? "000000"
  return "#" + (ref.tint ? applyTint(hex, ref.tint) : hex.toUpperCase())
}

/** Which of the theme's fonts a cell's font follows: headings (major) or body (minor). */
export type FontScheme = "major" | "minor"
