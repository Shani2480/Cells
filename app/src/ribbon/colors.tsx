import {
  OFFICE_THEME,
  SHEET_THEME_SLOTS,
  themeColorHex,
  type ShapeColor,
  type Theme,
  type ThemeColorRef,
} from "@cells/core"
import { useState } from "react"

/**
 * Excel's colour menus: Theme Colors (the workbook theme's ten colours with five lighter and darker steps each),
 * Standard Colors, and the Automatic / No Fill row. A theme colour is picked with its reference (slot and tint), so
 * cells and shapes keep following the theme when Page Layout › Themes changes it.
 */

/** The ten colours of the menu's first row, in SpreadsheetML's order (lt1, dk1, lt2, dk2, accent1–6). */
const ROLES = [
  "Background 1",
  "Text 1",
  "Background 2",
  "Text 2",
  "Accent 1",
  "Accent 2",
  "Accent 3",
  "Accent 4",
  "Accent 5",
  "Accent 6",
]

/** Office 2023's colour names, as Excel shows them for the default theme. */
const OFFICE_NAMES: Record<string, string> = {
  FFFFFF: "White",
  "000000": "Black",
  E8E8E8: "Light Gray",
  "0E2841": "Dark Blue",
  "156082": "Dark Teal",
  E97132: "Orange",
  "196B24": "Dark Green",
  "0F9ED5": "Turquoise",
  A02B93: "Purple",
  "4EA72E": "Green",
}

/** A plain name for a colour (Excel names theme colours by what they look like). */
export function colorName(hex: string): string {
  const up = hex.replace("#", "").toUpperCase()
  if (OFFICE_NAMES[up]) return OFFICE_NAMES[up]
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(up.slice(i, i + 2), 16) / 255)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const sat = max === min ? 0 : (max - min) / (1 - Math.abs(2 * l - 1))
  const shade = l < 0.3 ? "Dark " : l > 0.75 ? "Light " : ""
  if (sat < 0.15) return l > 0.95 ? "White" : l < 0.08 ? "Black" : `${shade}Gray`
  const d = max - min
  const h = 60 * (max === r ? ((g - b) / d + 6) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4)
  const hue =
    h < 15 || h >= 345
      ? "Red"
      : h < 40
        ? l < 0.4
          ? "Brown"
          : "Orange"
        : h < 55
          ? "Gold"
          : h < 70
            ? "Yellow"
            : h < 160
              ? "Green"
              : h < 190
                ? "Teal"
                : h < 205
                  ? "Turquoise"
                  : h < 255
                    ? "Blue"
                    : h < 290
                      ? "Purple"
                      : "Pink"
  return shade + hue
}

/** A swatch of the menu: its name, look and, for theme colours, the reference cells and shapes keep. */
export type Swatch = { name: string; hex: string; ref?: ThemeColorRef }

export const STANDARD_COLORS = [
  { name: "Dark Red", hex: "#C00000" },
  { name: "Red", hex: "#FF0000" },
  { name: "Orange", hex: "#FFC000" },
  { name: "Yellow", hex: "#FFFF00" },
  { name: "Light Green", hex: "#92D050" },
  { name: "Green", hex: "#00B050" },
  { name: "Light Blue", hex: "#00B0F0" },
  { name: "Blue", hex: "#0070C0" },
  { name: "Dark Blue", hex: "#002060" },
  { name: "Purple", hex: "#7030A0" },
]

/** The five variations Excel lists under a theme colour, as tint amounts (+ lighter, − darker). */
function variations(hex: string): number[] {
  const l = lightness(hex)
  if (l >= 0.99) return [-0.05, -0.15, -0.25, -0.35, -0.5]
  if (l <= 0.01) return [0.5, 0.35, 0.25, 0.15, 0.05]
  if (l > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9]
  if (l < 0.2) return [0.9, 0.75, 0.5, 0.25, 0.1]
  return [0.8, 0.6, 0.4, -0.25, -0.5]
}

function lightness(hex: string) {
  const n = parseInt(hex.replace("#", ""), 16)
  const [r, g, b] = [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255]
  return (Math.max(r, g, b) + Math.min(r, g, b)) / 2
}

/** Excel's tint amounts as it writes them (Excel stores 0.8 as 0.79998168889431442 and so on). */
const STORED_TINT: Record<string, number> = {
  "0.8": 0.7999816888943144,
  "0.6": 0.5999938962981048,
  "0.4": 0.3999755851924192,
  "0.5": 0.499984740745262,
  "0.35": 0.3499862666707358,
  "0.25": 0.249977111117893,
  "0.15": 0.1499679555650502,
  "0.05": 0.0499893185216834,
  "0.9": 0.8999908444471572,
  "0.75": 0.7499923703726921,
  "0.1": 0.0999786370433668,
  "-0.05": -0.0499893185216834,
  "-0.15": -0.1499984740745262,
  "-0.25": -0.249977111117893,
  "-0.35": -0.3499862666707358,
  "-0.5": -0.499984740745262,
  "-0.1": -0.0999786370433668,
  "-0.75": -0.7499923703726921,
  "-0.9": -0.8999908444471572,
}

function percent(t: number) {
  return t > 0 ? `Lighter ${Math.round(t * 100)}%` : `Darker ${Math.round(-t * 100)}%`
}

/** The theme grid, rows of ten: the theme's colours, then their five variations. */
export function themeGrid(theme: Theme = OFFICE_THEME): Swatch[][] {
  const first: Swatch[] = SHEET_THEME_SLOTS.slice(0, 10).map((slot, i) => {
    const hex = `#${theme.colors[slot].toUpperCase()}`
    return { name: `${colorName(hex)}, ${ROLES[i]}`, hex, ref: { theme: i } }
  })
  const rows = [first]
  for (let v = 0; v < 5; v++)
    rows.push(
      first.map((c, i) => {
        const t = variations(c.hex)[v]
        const ref = { theme: i, tint: STORED_TINT[String(t)] ?? t }
        return { name: `${c.name}, ${percent(t)}`, hex: themeColorHex(ref, theme), ref }
      }),
    )
  return rows
}

/** The grid of a theme, kept while the theme is the same. */
let cached: { theme: Theme; grid: Swatch[][] } | null = null
function gridOf(theme: Theme): Swatch[][] {
  if (!cached || cached.theme !== theme) cached = { theme, grid: themeGrid(theme) }
  return cached.grid
}

export function ColorPanel({
  noneLabel,
  onPick,
  current,
  theme = OFFICE_THEME,
}: {
  /** "Automatic" (font colour) or "No Fill"; without it the panel has no such row (sparkline colours). */
  noneLabel?: string
  /** The colour picked (null: Automatic / No Fill), with its theme reference for a theme colour. */
  onPick: (hex: string | null, ref?: ThemeColorRef) => void
  current?: string
  /** The workbook's theme, whose colours the first rows show. */
  theme?: Theme
}) {
  const [hover, setHover] = useState<string | null>(null)
  const GRID = gridOf(theme)
  const swatch = (c: Swatch, key: string, cls = "") => (
    <button
      key={key}
      type="button"
      role="menuitemradio"
      aria-checked={current?.toUpperCase() === c.hex}
      aria-label={c.name}
      title={c.name}
      className={`color-swatch ${cls}`}
      style={{ background: c.hex }}
      onMouseEnter={() => setHover(c.name)}
      onMouseLeave={() => setHover(null)}
      onClick={() => onPick(c.hex, c.ref)}
    />
  )
  return (
    <div className="color-panel" role="menu" aria-label={hover ?? undefined}>
      {noneLabel ? (
        <button type="button" role="menuitem" className="menu-row color-none" onClick={() => onPick(null)}>
          <span className="color-none-mark" data-none={noneLabel !== "Automatic" || undefined} />
          <span>{noneLabel}</span>
        </button>
      ) : null}
      <div className="color-heading">Theme Colors</div>
      <div className="color-grid">{GRID[0].map((c, i) => swatch(c, `t${i}`))}</div>
      <div className="color-grid is-tints">
        {GRID.slice(1).flatMap((row, r) => row.map((c, i) => swatch(c, `v${r}-${i}`, "is-tint")))}
      </div>
      <div className="color-heading">Standard Colors</div>
      <div className="color-grid">{STANDARD_COLORS.map((c, i) => swatch(c, `s${i}`))}</div>
    </div>
  )
}

/**
 * A colour picked from the menu as a shape keeps it: a theme colour (with its lighter or darker step, as DrawingML's
 * lumMod/lumOff) so it follows the workbook's theme, or a standard colour as RGB.
 */
export function shapeColorOf(hex: string, ref?: ThemeColorRef): ShapeColor {
  if (!ref) return { rgb: hex.replace("#", "").toUpperCase() }
  const slot = SHEET_THEME_SLOTS[ref.theme]
  const t = ref.tint ? Math.round(ref.tint * 100) / 100 : 0
  if (!t) return { theme: slot }
  return t > 0 ? { theme: slot, lumMod: 1 - t, lumOff: t } : { theme: slot, lumMod: 1 + t }
}
