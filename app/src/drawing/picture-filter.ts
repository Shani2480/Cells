/**
 * Picture Format › Color and Transparency, Office 2007's kind, drawn as an SVG filter on the picture (as Office applies
 * them when drawing): Set Transparent Color, Recolor (grayscale, black and white, sepia and the colour variations),
 * Office 2007's brightness and contrast (Washout), and transparency (the picture's opacity).
 */
import type { DuotoneColor, PictureAdjust, Theme } from "@cells/core"

const n = (v: number) => String(Math.round(v * 10000) / 10000)
/** Office's grey (Grayscale, Black and White, the colour variations): Rec. 709's weights on the picture's sRGB values,
 * as Excel's own pictures show. */
const GREY = "0.2126 0.7152 0.0722 0 0"

const toLinear = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4)
const toSrgb = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055)

/** A duotone colour as RGB (0–1): its theme or RGB colour, tinted or shaded (in linear light, as Office does), then
 * its saturation changed: scaled freely, each channel then clipped (Excel's colours, measured). */
export function duotoneRgb(c: DuotoneColor, theme: Theme): [number, number, number] {
  const hex =
    c.preset === "black"
      ? "000000"
      : c.preset === "white"
        ? "FFFFFF"
        : ((c.theme ? theme.colors[c.theme] : c.rgb) ?? "000000")
  let rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
  if (c.tint !== undefined || c.shade !== undefined) {
    rgb = rgb.map((v) => {
      let lin = toLinear(v)
      if (c.tint !== undefined) lin = lin * c.tint + (1 - c.tint)
      if (c.shade !== undefined) lin *= c.shade
      return toSrgb(Math.min(1, Math.max(0, lin)))
    })
  }
  if (c.satMod !== undefined) {
    const max = Math.max(...rgb)
    const min = Math.min(...rgb)
    const l = (max + min) / 2
    if (max > min) {
      rgb = rgb.map((v) => Math.min(1, Math.max(0, l + (v - l) * c.satMod!)))
    }
  }
  return rgb as [number, number, number]
}

/** "0 0 … 1 … 0": 1 only at `value` (0–255), for a discrete transfer that finds one channel value exactly. */
function only(value: number): string {
  const out = new Array<string>(256).fill("0")
  out[value] = "1"
  return out.join(" ")
}

/**
 * The SVG filter for a picture's Office 2007 adjustments (`id` names it), or "" when it has none. Transparency is not
 * in it: it's the picture's opacity (`pictureOpacity`).
 */
export function pictureFilter(id: string, a: PictureAdjust | undefined, theme: Theme): string {
  if (!a) return ""
  const steps: string[] = []
  if (a.transparentColor) {
    // Pixels of exactly that colour become see-through: each channel's match is 1, and all three must match.
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(a.transparentColor!.slice(i, i + 2), 16))
    steps.push(
      `<feComponentTransfer in="SourceGraphic" result="hit"><feFuncR type="discrete" tableValues="${only(r)}"/>` +
        `<feFuncG type="discrete" tableValues="${only(g)}"/><feFuncB type="discrete" tableValues="${only(b)}"/></feComponentTransfer>` +
        `<feColorMatrix in="hit" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 1 1 0 -2" result="mask"/>` +
        `<feComposite in="SourceGraphic" in2="mask" operator="out" result="s0"/>`,
    )
  }
  const input = () => (steps.length ? `in="s${steps.length - 1}" ` : `in="SourceGraphic" `)
  const result = () => `result="s${steps.length}"`
  const r = a.recolor
  if (r) {
    const grey = `<feColorMatrix ${input()}type="matrix" values="${GREY} ${GREY} ${GREY} 0 0 0 1 0"`
    if (r.kind === "grayscale") steps.push(`${grey} ${result()}/>`)
    else if (r.kind === "blackWhite") {
      const table = Array.from({ length: 100 }, (_, i) => (i / 100 >= r.threshold ? "1" : "0")).join(" ")
      const f = (c: string) => `<feFunc${c} type="discrete" tableValues="${table}"/>`
      steps.push(
        `${grey} result="g${steps.length}"/><feComponentTransfer in="g${steps.length}" ${result()}>${f("R")}${f("G")}${f("B")}</feComponentTransfer>`,
      )
    } else {
      const [dark, light] = r.colors.map((c) => duotoneRgb(c, theme))
      const f = (c: string, i: number) => `<feFunc${c} type="table" tableValues="${n(dark[i])} ${n(light[i])}"/>`
      steps.push(
        `${grey} result="g${steps.length}"/><feComponentTransfer in="g${steps.length}" ${result()}>${f("R", 0)}${f("G", 1)}${f("B", 2)}</feComponentTransfer>`,
      )
    }
  }
  if (a.lum && (a.lum.bright || a.lum.contrast)) {
    // Office 2007's brightness and contrast, one straight line per channel (fitted to Excel's pictures): contrast is
    // the slope about the middle grey (1 + c below 0, 1 / (1 - c) above), and brightness lifts it by b × (1 + slope) / 2.
    const c = a.lum.contrast
    const slope = c < 0 ? 1 + c : 1 / Math.max(0.001, 1 - c)
    const intercept = 0.5 * (1 - slope) + (a.lum.bright * (1 + slope)) / 2
    const f = (ch: string) => `<feFunc${ch} type="linear" slope="${n(slope)}" intercept="${n(intercept)}"/>`
    steps.push(`<feComponentTransfer ${input()}${result()}>${f("R")}${f("G")}${f("B")}</feComponentTransfer>`)
  }
  if (!steps.length) return ""
  return `<filter id="${id}" x="0" y="0" width="1" height="1" color-interpolation-filters="sRGB">${steps.join("")}</filter>`
}

/** The picture's opacity (Transparency), or null when it's opaque. */
export function pictureOpacity(a: PictureAdjust | undefined): number | null {
  return a?.transparency ? Math.max(0, 1 - a.transparency) : null
}
