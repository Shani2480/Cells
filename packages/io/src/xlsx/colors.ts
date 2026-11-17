import { applyTint } from "@cells/core"
import { INDEXED_COLORS } from "./constants"

/**
 * Colours in SpreadsheetML: `rgb` (ARGB hex), `theme` (index into the theme's colour scheme, with an optional
 * `tint`), `indexed` (the legacy palette) or `auto`. Cells keeps "#rrggbb"; theme and indexed colours resolve here.
 */

/** The theme's colour scheme in the order SpreadsheetML's `theme` index uses: lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink. */
export type ThemeColors = string[]

/** Office's default theme (2023+, the one Excel 365 writes for new workbooks). */
export const DEFAULT_THEME_COLORS: ThemeColors = [
  "FFFFFF", "000000", "E8E8E8", "0E2841", "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "467886", "96607D",
] // prettier-ignore

/** Reads the clrScheme of a theme part: dk1 lt1 dk2 lt2 accent1–6 hlink folHlink, swapped into SpreadsheetML order. */
export function themeColorsFrom(scheme: Record<string, string>): ThemeColors {
  const order = [
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
  return order.map((k, i) => scheme[k] ?? DEFAULT_THEME_COLORS[i])
}

export type ColorAttrs = { rgb?: string; theme?: string; tint?: string; indexed?: string; auto?: string }

/** A colour element's attributes as "#rrggbb", or null for automatic / unknown. */
export { applyTint }

export function resolveColor(attrs: ColorAttrs, theme: ThemeColors, palette: string[] = INDEXED_COLORS): string | null {
  let hex: string | null = null
  if (attrs.rgb) hex = attrs.rgb.length === 8 ? attrs.rgb.slice(2) : attrs.rgb
  else if (attrs.theme !== undefined) hex = theme[Number(attrs.theme)] ?? null
  else if (attrs.indexed !== undefined) {
    const i = Number(attrs.indexed)
    hex = palette[i] ?? null // 64/65: system colours (automatic)
  }
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return null
  const tint = attrs.tint ? Number(attrs.tint) : 0
  return "#" + (tint ? applyTint(hex, tint) : hex).toUpperCase()
}

/** "#rrggbb" as SpreadsheetML's ARGB. */
export const argb = (hex: string) => "FF" + hex.replace("#", "").toUpperCase()
