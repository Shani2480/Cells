import type { ShapeFill, TextEffects } from "@cells/core"

/** Insert › WordArt: Excel's styles (a selection), as letter fill and effects in theme colours. */
export const WORDART_STYLES: { name: string; fill?: ShapeFill; effects?: TextEffects }[] = [
  {
    name: "Fill: Black, Text color 1; Shadow",
    fill: { kind: "solid", color: { theme: "dk1" } },
    effects: { shadow: true },
  },
  {
    name: "Fill: Blue, Accent color 1; Shadow",
    fill: { kind: "solid", color: { theme: "accent1" } },
    effects: { shadow: true },
  },
  {
    name: "Fill: Orange, Accent color 2; Outline: Orange, Accent color 2",
    fill: { kind: "solid", color: { theme: "accent2", lumMod: 0.6, lumOff: 0.4 } },
    effects: { outline: { color: { theme: "accent2" }, width: 1 } },
  },
  {
    name: "Fill: White; Outline: Blue, Accent color 1; Glow: Blue, Accent color 1",
    fill: { kind: "solid", color: { theme: "lt1" } },
    effects: { outline: { color: { theme: "accent1" }, width: 1 }, glow: { theme: "accent1", alpha: 0.6 } },
  },
  {
    name: "Fill: Dark Green, Accent color 3; Reflection",
    fill: { kind: "solid", color: { theme: "accent3" } },
    effects: { reflection: true },
  },
  {
    name: "Fill: White; Outline: Black, Text color 1",
    fill: { kind: "solid", color: { theme: "lt1" } },
    effects: { outline: { color: { theme: "dk1" }, width: 1 }, shadow: true },
  },
  {
    name: "Fill: Purple, Accent color 5; Outline: White",
    fill: { kind: "solid", color: { theme: "accent5" } },
    effects: { outline: { color: { theme: "lt1" }, width: 0.75 }, glow: { theme: "accent5", alpha: 0.4 } },
  },
  {
    name: "Fill: Gray, Background color 2; Inner Shadow",
    fill: { kind: "solid", color: { theme: "lt2", lumMod: 0.5 } },
    effects: { shadow: true },
  },
]
