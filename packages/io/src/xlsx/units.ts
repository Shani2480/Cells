/**
 * Sizes as Excel stores them against Cells' pixels. Column widths are in characters of the default font's widest
 * digit (7 px for Aptos Narrow / Calibri 11), plus 5 px of padding; row heights are in points.
 */
export const MAX_DIGIT_WIDTH = 7

/** A file's column width (characters, padding included) in pixels, as Excel draws it at 100%. */
export function colWidthToPx(width: number, digit = MAX_DIGIT_WIDTH): number {
  if (!(width > 0)) return 0
  return Math.trunc(((256 * width + Math.trunc(128 / digit)) / 256) * digit)
}

/** Pixels as a file's column width: Excel's own rounding to 1/256 of a character. */
export function pxToColWidth(px: number, digit = MAX_DIGIT_WIDTH): number {
  return px <= 0 ? 0 : Math.trunc((px / digit) * 256) / 256
}

export const pointsToPx = (pt: number) => Math.round((pt * 4) / 3)
export const pxToPoints = (px: number) => Math.round(px * 0.75 * 100) / 100
