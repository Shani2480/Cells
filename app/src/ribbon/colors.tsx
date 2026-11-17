import type { ShapeColor } from "@cells/core"
import { useState } from "react"

/**
 * Excel's colour menus: Theme Colors (ten theme colours with five tints and shades each), Standard Colors, and the
 * Automatic / No Fill row. Theme colours are the Office 2023 theme; tints use Excel's luminance rule.
 */

export const THEME_COLORS = [
  { name: "White, Background 1", hex: "#FFFFFF" },
  { name: "Black, Text 1", hex: "#000000" },
  { name: "Light Gray, Background 2", hex: "#E8E8E8" },
  { name: "Dark Blue, Text 2", hex: "#0E2841" },
  { name: "Blue, Accent 1", hex: "#156082" },
  { name: "Orange, Accent 2", hex: "#E97132" },
  { name: "Dark Green, Accent 3", hex: "#196B24" },
  { name: "Turquoise, Accent 4", hex: "#0F9ED5" },
  { name: "Purple, Accent 5", hex: "#A02B93" },
  { name: "Green, Accent 6", hex: "#4EA72E" },
]

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
  const [r, g, b] = rgb(hex)
  return (Math.max(r, g, b) + Math.min(r, g, b)) / 2
}

function rgb(hex: string): [number, number, number] {
  const n = parseInt(hex.slice(1), 16)
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255]
}

/** Excel's tint: lighter moves luminance toward white by `t`, darker scales it by 1 + t (in HSL). */
export function tint(hex: string, t: number): string {
  const [r, g, b] = rgb(hex)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let l = (max + min) / 2
  const d = max - min
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1))
  let h = 0
  if (d) {
    if (max === r) h = ((g - b) / d) % 6
    else if (max === g) h = (b - r) / d + 2
    else h = (r - g) / d + 4
    h = (h * 60 + 360) % 360
  }
  l = t < 0 ? l * (1 + t) : l * (1 - t) + t
  const c = (1 - Math.abs(2 * l - 1)) * s
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1))
  const m = l - c / 2
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
  const out = (v: number) =>
    Math.round(Math.min(1, Math.max(0, v + m)) * 255)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase()
  return `#${out(r1)}${out(g1)}${out(b1)}`
}

function percent(t: number) {
  return t > 0 ? `Lighter ${Math.round(t * 100)}%` : `Darker ${Math.round(-t * 100)}%`
}

/** The theme grid, rows of ten: the colours, then their five variations. */
export function themeGrid(): { name: string; hex: string }[][] {
  const rows = [THEME_COLORS]
  for (let i = 0; i < 5; i++)
    rows.push(
      THEME_COLORS.map((c) => {
        const t = variations(c.hex)[i]
        return { name: `${c.name}, ${percent(t)}`, hex: tint(c.hex, t) }
      }),
    )
  return rows
}

const GRID = themeGrid()

export function ColorPanel({
  noneLabel,
  onPick,
  current,
}: {
  /** "Automatic" (font colour) or "No Fill". */
  noneLabel: string
  onPick: (hex: string | null) => void
  current?: string
}) {
  const [hover, setHover] = useState<string | null>(null)
  const swatch = (c: { name: string; hex: string }, key: string, cls = "") => (
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
      onClick={() => onPick(c.hex)}
    />
  )
  return (
    <div className="color-panel" role="menu" aria-label={hover ?? undefined}>
      <button type="button" role="menuitem" className="menu-row color-none" onClick={() => onPick(null)}>
        <span className="color-none-mark" data-none={noneLabel !== "Automatic" || undefined} />
        <span>{noneLabel}</span>
      </button>
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

const SLOTS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const

/**
 * A colour picked from the menu as a shape keeps it: a theme colour (with its lighter or darker step, as DrawingML's
 * lumMod/lumOff) so it follows the workbook's theme, or a standard colour as RGB.
 */
export function shapeColorOf(hex: string): ShapeColor {
  const up = hex.toUpperCase()
  for (let r = 0; r < GRID.length; r++) {
    const i = GRID[r].findIndex((c) => c.hex.toUpperCase() === up)
    if (i < 0) continue
    if (r === 0) return { theme: SLOTS[i] }
    const t = variations(THEME_COLORS[i].hex)[r - 1]
    return t > 0 ? { theme: SLOTS[i], lumMod: 1 - t, lumOff: t } : { theme: SLOTS[i], lumMod: 1 + t }
  }
  return { rgb: up.replace("#", "") }
}
