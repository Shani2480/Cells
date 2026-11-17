import { themeColorHex, type FontScheme, type Theme, type ThemeColorRef } from "./theme"

/**
 * Cell formats, interned: a cell, row or column stores a style id, the workbook keeps one record per distinct format.
 * Id 0 is the default (General, no direct formatting): a cell with style 0 shows its row's format, else its
 * column's. Id 1 is also the default format, but set on purpose (Clear Formats on a cell of a formatted column), so
 * it does not fall back to the row or column.
 *
 * Fields follow Excel's cell format record (xf); an absent field is Excel's default.
 */

export type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "hair"
  | "dotted"
  | "dashed"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"
  | "double"

/**
 * A border line; no colour means Automatic (the window text colour, black). `theme`: the colour is the theme's (and
 * `color` its look under the current theme).
 */
export type Border = { style: BorderStyle; color?: string; theme?: ThemeColorRef }

export type HAlign = "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed"
export type VAlign = "top" | "center" | "bottom" | "justify" | "distributed"
export type Underline = "single" | "double" | "singleAccounting" | "doubleAccounting"

export type CellStyle = {
  numberFormat?: string
  // Font
  fontName?: string
  /** Points. */
  fontSize?: number
  bold?: true
  italic?: true
  underline?: Underline
  strike?: true
  vertAlign?: "superscript" | "subscript"
  /** "#rrggbb"; absent is Automatic. */
  fontColor?: string
  /** The font colour is this theme colour (`fontColor` holds how it looks with the workbook's theme). */
  fontThemeColor?: ThemeColorRef
  /** The font is the theme's headings or body font (`fontName` holds the current theme's). */
  fontScheme?: FontScheme
  // Fill
  /** Solid fill, "#rrggbb"; absent is No Fill. */
  fill?: string
  /** The fill is this theme colour (`fill` holds how it looks with the workbook's theme). */
  fillThemeColor?: ThemeColorRef
  // Alignment
  hAlign?: HAlign
  /** Absent is bottom, Excel's default. */
  vAlign?: VAlign
  indent?: number
  wrap?: true
  shrink?: true
  /** Excel's encoding: 1–90 counterclockwise, 91–180 clockwise (91 = 1° down), 255 stacked vertical text. */
  rotation?: number
  // Borders
  borderTop?: Border
  borderBottom?: Border
  borderLeft?: Border
  borderRight?: Border
  borderDiagonal?: Border
  diagonalUp?: true
  diagonalDown?: true
  // Protection (applies when the sheet is protected)
  /** Cells are locked by default; false unlocks. */
  locked?: false
  hiddenFormula?: true
  /** The cell style applied from the gallery ("Good", "Heading 1"…), Normal when absent. */
  cellStyle?: string
}

export type StyleKey = keyof CellStyle

/** Changes to a format: a value sets the field, null removes it (back to Excel's default). */
export type StylePatch = { [K in StyleKey]?: CellStyle[K] | null }

export const DEFAULT_STYLE_ID = 0
/** The default format, set explicitly: it does not fall back to the row's or column's format. */
export const EXPLICIT_DEFAULT_ID = 1

export class StyleTable {
  private readonly styles: CellStyle[] = [Object.freeze({}), Object.freeze({})]
  private readonly index = new Map<string, number>([[keyOf({}), 0]])
  /** (base id, patch) → result id: formatting 100,000 cells with a few distinct formats derives a few times. */
  private readonly patched = new Map<string, number>()

  get(id: number): CellStyle {
    return this.styles[id] ?? this.styles[0]
  }

  /** The id of a style with these fields, adding it if new. */
  intern(style: CellStyle): number {
    const key = keyOf(style)
    const found = this.index.get(key)
    if (found !== undefined) {
      return found
    }
    const id = this.styles.length
    this.styles.push(Object.freeze({ ...style }))
    this.index.set(key, id)
    return id
  }

  /** The id of `base` with `patch` applied (null fields removed). */
  derive(base: number, patch: StylePatch): number {
    // A field the patch leaves undefined is not touched; null removes it.
    const merged: Record<string, unknown> = { ...this.get(base) }
    for (const [key, value] of Object.entries(patch)) if (value !== undefined) merged[key] = value
    // A colour or font set on its own is no longer the theme's.
    for (const [field, link] of THEME_LINKS)
      if (patch[field] !== undefined && patch[link] === undefined) delete merged[link]
    for (const key of Object.keys(merged)) {
      if (merged[key] === undefined || merged[key] === null) {
        delete merged[key]
      }
    }
    return this.intern(merged as CellStyle)
  }

  /** As derive, cached by the patch's key (from `patchKey`). */
  deriveCached(base: number, patch: StylePatch, key: string): number {
    const cacheKey = `${base}|${key}`
    let id = this.patched.get(cacheKey)
    if (id === undefined) {
      id = this.derive(base, patch)
      this.patched.set(cacheKey, id)
    }
    return id
  }

  /**
   * Page Layout › Themes: every format that uses a theme colour or font takes the look of `theme` (ids stay the same,
   * so cells keep their formats).
   */
  retheme(theme: Theme) {
    let changed = false
    for (let id = 2; id < this.styles.length; id++) {
      const next = themed(this.styles[id], theme)
      if (next) {
        this.styles[id] = Object.freeze(next)
        changed = true
      }
    }
    if (!changed) return
    this.index.clear()
    this.styles.forEach((style, id) => {
      const key = keyOf(style)
      if (id !== 1 && !this.index.has(key)) this.index.set(key, id)
    })
    this.patched.clear()
  }

  /** Whether the id stands for Excel's default format (0, or 1 set on purpose). */
  isDefault(id: number) {
    return id === DEFAULT_STYLE_ID || id === EXPLICIT_DEFAULT_ID
  }

  get size() {
    return this.styles.length
  }
}

/** Fields that a theme reference goes with: setting the field alone drops the reference. */
const THEME_LINKS: [StyleKey, StyleKey][] = [
  ["fontColor", "fontThemeColor"],
  ["fill", "fillThemeColor"],
  ["fontName", "fontScheme"],
]

const BORDER_KEYS = ["borderTop", "borderBottom", "borderLeft", "borderRight", "borderDiagonal"] as const

/** A format with its theme colours and fonts as `theme` shows them; null when nothing changes. */
export function themed(style: CellStyle, theme: Theme): CellStyle | null {
  let out: CellStyle | null = null
  const set = <K extends StyleKey>(key: K, value: CellStyle[K]) => {
    if (style[key] === value) return
    out ??= { ...style }
    out[key] = value
  }
  if (style.fontThemeColor) set("fontColor", themeColorHex(style.fontThemeColor, theme))
  if (style.fillThemeColor) set("fill", themeColorHex(style.fillThemeColor, theme))
  if (style.fontScheme) set("fontName", theme.fonts[style.fontScheme])
  for (const key of BORDER_KEYS) {
    const b = style[key]
    if (b?.theme && b.color !== themeColorHex(b.theme, theme)) set(key, { ...b, color: themeColorHex(b.theme, theme) })
  }
  return out
}

/** A stable text key for a patch, for caching. */
export function patchKey(patch: StylePatch): string {
  return JSON.stringify(
    Object.keys(patch)
      .filter((key) => patch[key as StyleKey] !== undefined)
      .sort()
      .map((key) => [key, patch[key as StyleKey]]),
  )
}

function keyOf(style: CellStyle) {
  return patchKey(style)
}

/** The patch that removes every field: Clear Formats. */
export const CLEAR_ALL: StylePatch = {
  numberFormat: null,
  fontName: null,
  fontSize: null,
  bold: null,
  italic: null,
  underline: null,
  strike: null,
  vertAlign: null,
  fontColor: null,
  fontThemeColor: null,
  fontScheme: null,
  fill: null,
  fillThemeColor: null,
  hAlign: null,
  vAlign: null,
  indent: null,
  wrap: null,
  shrink: null,
  rotation: null,
  borderTop: null,
  borderBottom: null,
  borderLeft: null,
  borderRight: null,
  borderDiagonal: null,
  diagonalUp: null,
  diagonalDown: null,
  locked: null,
  hiddenFormula: null,
  cellStyle: null,
}
