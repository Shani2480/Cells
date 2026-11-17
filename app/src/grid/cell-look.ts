import {
  formatMark,
  formatValue,
  formatWithCode,
  generalAlign,
  hasFormatMarks,
  isError,
  type Border,
  type CellStyle,
  type CellValue,
  type StyleTable,
} from "@cells/core"
import { CELL_FONT_FAMILY } from "./theme"

/** Text width in canvas pixels; `font` defaults to the sheet's body font. */
export type Measure = (text: string, font?: string) => number

/** Excel's default body font size, points. */
export const DEFAULT_FONT_PT = 11
const PX_PER_PT = 96 / 72

/** A style as the painter needs it, computed once per style id. */
export type CellLook = {
  style: CellStyle
  /** Canvas font shorthand. */
  font: string
  fontPx: number
  /** Explicit colour; undefined is Automatic (the theme's text colour). */
  color?: string
  fill?: string
  underline?: CellStyle["underline"]
  strike: boolean
  hAlign: NonNullable<CellStyle["hAlign"]>
  vAlign: NonNullable<CellStyle["vAlign"]>
  /** Pixels of indent. */
  indent: number
  wrap: boolean
  shrink: boolean
  /** Degrees, counterclockwise positive; null for none; "stacked" for vertical text. */
  rotation: number | null | "stacked"
  top?: Border
  bottom?: Border
  left?: Border
  right?: Border
  diagonal?: Border
  diagonalUp: boolean
  diagonalDown: boolean
}

/** Pixels per indent level at 100% (Excel indents by about three characters' worth of the default font). */
export const INDENT_PX = 9

/** The workbook theme's body font, which cells without a font of their own use (Excel's Normal style). */
let bodyFont = "Aptos Narrow"

/** Sets the body font (Page Layout › Themes); looks made with another are dropped. */
export function setBodyFont(name: string) {
  if (name === bodyFont) return
  bodyFont = name
  looks = new WeakMap()
}

export function fontFor(style: CellStyle, scale = 1): { font: string; px: number } {
  const px = (style.fontSize ?? DEFAULT_FONT_PT) * PX_PER_PT * scale
  const name = style.fontName ?? bodyFont
  const family = name && name !== "Aptos Narrow" ? `"${name.replace(/"/g, "")}", ${CELL_FONT_FAMILY}` : CELL_FONT_FAMILY
  const font = `${style.italic ? "italic " : ""}${style.bold ? "bold " : ""}${round2(px)}px ${family}`
  return { font, px }
}

const round2 = (n: number) => Math.round(n * 100) / 100

function rotationOf(r: number | undefined): CellLook["rotation"] {
  if (r === undefined || r === 0) return null
  if (r === 255) return "stacked"
  return r <= 90 ? r : -(r - 90)
}

let looks = new WeakMap<StyleTable, Map<number, CellLook>>()

export function lookOf(styles: StyleTable, id: number): CellLook {
  let table = looks.get(styles)
  if (!table) looks.set(styles, (table = new Map()))
  let look = table.get(id)
  if (look) return look
  const style = styles.get(id)
  const { font, px } = fontFor(style)
  look = {
    style,
    font,
    fontPx: px,
    color: style.fontColor,
    fill: style.fill,
    underline: style.underline,
    strike: style.strike === true,
    hAlign: style.hAlign ?? "general",
    vAlign: style.vAlign ?? "bottom",
    indent: (style.indent ?? 0) * INDENT_PX,
    wrap: style.wrap === true,
    shrink: style.shrink === true,
    rotation: rotationOf(style.rotation),
    top: style.borderTop,
    bottom: style.borderBottom,
    left: style.borderLeft,
    right: style.borderRight,
    diagonal: style.borderDiagonal,
    diagonalUp: style.diagonalUp === true,
    diagonalDown: style.diagonalDown === true,
  }
  table.set(id, look)
  return look
}

/** What a cell shows: its text (with format marks), the format's colour, and how it aligns by default. */
export type Shown = {
  text: string
  /** Colour from the number format ([Red]). */
  color?: string
  /** Numbers (and dates, logicals, errors) never run into the next cell, and show #### when too wide. */
  numeric: boolean
  align: "left" | "center" | "right"
}

/**
 * The text of a value in a column `width` pixels wide. Numbers that don't fit become #### (General first shortens its
 * decimals, as Excel does).
 */
export function shownText(value: CellValue, look: CellLook, width: number, pad: number, measure: Measure): Shown {
  const code = look.style.numberFormat
  const generalAlignment = generalAlign(value)
  const align: Shown["align"] =
    look.hAlign === "left" || look.hAlign === "right" || look.hAlign === "center"
      ? look.hAlign
      : look.hAlign === "centerContinuous" || look.hAlign === "distributed"
        ? "center"
        : look.hAlign === "fill" || look.hAlign === "justify"
          ? "left"
          : generalAlignment
  if (typeof value === "string") {
    if (code && code !== "General" && code.includes("@")) {
      const r = formatWithCode(value, code, { cell: true })
      return { text: r.text, color: r.color, numeric: false, align }
    }
    return { text: value, numeric: false, align }
  }
  if (typeof value === "boolean" || isError(value)) {
    return { text: formatValue(value) ?? "", numeric: true, align: look.hAlign === "general" ? "center" : align }
  }
  const room = width - 2 * pad - look.indent
  if (code && code !== "General") {
    const r = formatWithCode(value, code, { cell: true })
    const text = r.text
    if (textWidth(text, look.font, measure) > room + pad)
      return { text: hashes(width, look, measure), numeric: true, align }
    return { text, color: r.color, numeric: true, align }
  }
  const digit = measure("0", look.font) || 7
  const maxChars = Math.max(1, Math.floor((room + 1) / digit))
  let text = formatValue(value, undefined, maxChars)
  if (text !== null && measure(text, look.font) > room + pad)
    text = formatValue(value, undefined, Math.max(1, maxChars - 1))
  if (text === null || measure(text, look.font) > width - 2) text = hashes(width, look, measure)
  return { text, numeric: true, align }
}

function hashes(width: number, look: CellLook, measure: Measure) {
  return "#".repeat(Math.max(1, Math.floor((width - 2) / (measure("#", look.font) || 7))))
}

/** Width of cell-mode text: gaps count as their character's width, fills as nothing (they take what is left). */
export function textWidth(text: string, font: string, measure: Measure): number {
  if (!hasFormatMarks(text)) return measure(text, font)
  let width = 0
  let run = ""
  for (const c of text) {
    const mark = formatMark(c)
    if (!mark) {
      run += c
      continue
    }
    width += measure(run, font)
    run = ""
    if (mark.kind === "pad") width += measure(mark.ch, font)
  }
  return width + measure(run, font)
}

/**
 * Lays out cell-mode text into `width` pixels: plain runs with their x offsets, gaps left blank, and the fill
 * character repeated to take the room that is left (Accounting's "$" at the left edge, the amount at the right).
 */
export function layoutMarks(
  text: string,
  font: string,
  measure: Measure,
  width: number,
): { runs: { text: string; x: number }[]; width: number } {
  const parts: ({ kind: "text"; text: string } | { kind: "pad"; w: number } | { kind: "fill"; ch: string })[] = []
  let run = ""
  for (const c of text) {
    const mark = formatMark(c)
    if (!mark) {
      run += c
      continue
    }
    if (run) parts.push({ kind: "text", text: run })
    run = ""
    parts.push(mark.kind === "pad" ? { kind: "pad", w: measure(mark.ch, font) } : { kind: "fill", ch: mark.ch })
  }
  if (run) parts.push({ kind: "text", text: run })
  const fixed = parts.reduce((w, p) => w + (p.kind === "text" ? measure(p.text, font) : p.kind === "pad" ? p.w : 0), 0)
  const fill = parts.find((p) => p.kind === "fill")
  const spare = Math.max(0, width - fixed)
  const runs: { text: string; x: number }[] = []
  let x = 0
  let filled = false
  for (const p of parts) {
    if (p.kind === "text") {
      runs.push({ text: p.text, x })
      x += measure(p.text, font)
    } else if (p.kind === "pad") {
      x += p.w
    } else if (!filled) {
      // Only the first fill counts, as in Excel.
      filled = true
      const one = measure(p.ch, font)
      if (p.ch !== " " && one > 0) {
        const n = Math.floor(spare / one)
        if (n > 0) runs.push({ text: p.ch.repeat(n), x })
      }
      x += spare
    }
  }
  return { runs, width: fill ? Math.max(fixed, width) : fixed }
}

/** Excel's 56-colour palette for [ColorN] in number formats. */
const PALETTE = [
  "#000000",
  "#FFFFFF",
  "#FF0000",
  "#00FF00",
  "#0000FF",
  "#FFFF00",
  "#FF00FF",
  "#00FFFF",
  "#800000",
  "#008000",
  "#000080",
  "#808000",
  "#800080",
  "#008080",
  "#C0C0C0",
  "#808080",
  "#9999FF",
  "#993366",
  "#FFFFCC",
  "#CCFFFF",
  "#660066",
  "#FF8080",
  "#0066CC",
  "#CCCCFF",
  "#000080",
  "#FF00FF",
  "#FFFF00",
  "#00FFFF",
  "#800080",
  "#800000",
  "#008080",
  "#0000FF",
  "#00CCFF",
  "#CCFFFF",
  "#CCFFCC",
  "#FFFF99",
  "#99CCFF",
  "#FF99CC",
  "#CC99FF",
  "#FFCC99",
  "#3366FF",
  "#33CCCC",
  "#99CC00",
  "#FFCC00",
  "#FF9900",
  "#FF6600",
  "#666699",
  "#969696",
  "#003366",
  "#339966",
  "#003300",
  "#333300",
  "#993300",
  "#993366",
  "#333399",
  "#333333",
]

/** A number format's colour as CSS: hex as given, [ColorN] from the palette. */
export function formatColor(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^color\s*(\d+)$/i.exec(color)
  if (m) return PALETTE[Number(m[1]) - 1]
  return color
}

/** Whether black text reads better than white on `fill` (for Automatic text on a filled cell in dark mode). */
export function prefersDarkText(fill: string): boolean {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(fill)
  if (!m) return true
  const [r, g, b] = [m[1], m[2], m[3]].map((h) => parseInt(h, 16) / 255)
  const lin = (c: number) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4)
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b) > 0.4
}

/**
 * A colour as dark mode shows it: the same hue with its lightness turned around, so light fills become dark and dark
 * text becomes light (as Excel's dark cells do). Greys and colours in the middle barely move.
 */
export function darkVariant(color: string): string {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(color)
  if (!m) return color
  const [r, g, b] = [m[1], m[2], m[3]].map((h) => parseInt(h, 16) / 255)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  let h = 0
  const sat = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1))
  if (d) {
    if (max === r) h = ((g - b) / d) % 6
    else if (max === g) h = (b - r) / d + 2
    else h = (r - g) / d + 4
    h *= 60
    if (h < 0) h += 360
  }
  // Dark backgrounds aren't pure black: keep the darkest results a little above it.
  const l2 = 0.08 + (1 - l) * 0.84
  const c = (1 - Math.abs(2 * l2 - 1)) * sat
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1))
  const k = l2 - c / 2
  const [r1, g1, b1] =
    h < 60
      ? [c, x, 0]
      : h < 120
        ? [x, c, 0]
        : h < 180
          ? [0, c, x]
          : h < 240
            ? [0, x, c]
            : h < 300
              ? [x, 0, c]
              : [c, 0, x]
  const hex = (v: number) =>
    Math.round(Math.min(1, Math.max(0, v + k)) * 255)
      .toString(16)
      .padStart(2, "0")
  return `#${hex(r1)}${hex(g1)}${hex(b1)}`
}
