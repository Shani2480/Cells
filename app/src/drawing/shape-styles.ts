import type { ShapeColor, ShapeFill, ShapeLine, SheetShape, ThemeSlot } from "@cells/core"

/**
 * Shape Format › Shape Styles: Excel's theme styles, six rows in dark 1 and the six accents. The last rows take the
 * theme's effects (Page Layout › Effects), as Excel's do; Cells fills them with the solid colour (Excel's gradients of
 * the theme fills aren't drawn).
 */
export type ShapeStyle = {
  name: string
  fill: ShapeFill
  line: ShapeLine
  text: ShapeColor
  effect?: 1 | 2 | 3
}

export const STYLE_COLORS: { slot: ThemeSlot; name: string }[] = [
  { slot: "dk1", name: "Dark 1" },
  { slot: "accent1", name: "Accent 1" },
  { slot: "accent2", name: "Accent 2" },
  { slot: "accent3", name: "Accent 3" },
  { slot: "accent4", name: "Accent 4" },
  { slot: "accent5", name: "Accent 5" },
  { slot: "accent6", name: "Accent 6" },
]

const ROWS: { name: string; style: (c: ThemeSlot) => Omit<ShapeStyle, "name"> }[] = [
  {
    name: "Colored Outline",
    style: (c) => ({ fill: solid({ theme: "lt1" }), line: { color: { theme: c }, width: 1 }, text: { theme: c } }),
  },
  {
    name: "Colored Fill",
    style: (c) => ({
      fill: solid({ theme: c }),
      line: { color: { theme: c, lumMod: 0.5 }, width: 1 },
      text: { theme: "lt1" },
    }),
  },
  {
    name: "Light 1 Outline, Colored Fill",
    style: (c) => ({
      fill: solid({ theme: c }),
      line: { color: { theme: "lt1" }, width: 1.5 },
      text: { theme: "lt1" },
      effect: 1,
    }),
  },
  {
    name: "Subtle Effect",
    style: (c) => ({
      fill: solid({ theme: c, lumMod: 0.2, lumOff: 0.8 }),
      line: { color: { theme: c }, width: 0.5 },
      text: c === "dk1" ? { theme: "dk1" } : { theme: c, lumMod: 0.5 },
      effect: 1,
    }),
  },
  {
    name: "Moderate Effect",
    style: (c) => ({
      fill: solid({ theme: c, lumMod: 0.75 }),
      line: { color: null, width: 1 },
      text: { theme: "lt1" },
      effect: 2,
    }),
  },
  {
    name: "Intense Effect",
    style: (c) => ({ fill: solid({ theme: c }), line: { color: null, width: 1 }, text: { theme: "lt1" }, effect: 3 }),
  },
]

function solid(color: ShapeColor): ShapeFill {
  return { kind: "solid", color }
}

/** The gallery: rows of styles, each "<row> - <colour>" as Excel names them. */
export const SHAPE_STYLES: ShapeStyle[][] = ROWS.map((row) =>
  STYLE_COLORS.map(({ slot, name }) => ({ name: `${row.name} - ${name}`, ...row.style(slot) })),
)

/** A shape with a style applied: its fill, outline, text colour and theme effect (its own shadow goes). */
export function applyShapeStyle(o: SheetShape, s: ShapeStyle) {
  o.fill = structuredClone(s.fill)
  o.line = { ...o.line, color: s.line.color && { ...s.line.color }, width: s.line.width }
  if (o.text) for (const p of o.text.paragraphs) for (const r of p.runs) r.color = { ...s.text }
  delete o.shadow
  if (s.effect) o.themeEffect = s.effect
  else delete o.themeEffect
}
