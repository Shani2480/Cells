import { embedSvg } from "./embed"
import { signatureSvg } from "./signature"
import { layoutEquation, type MathBox } from "./math"
import {
  colorCss,
  effectLooks,
  isChart,
  type DrawingObject,
  type EffectLook,
  type InkStroke,
  type ShapeColor,
  type ShapeFill,
  type ShapeLine,
  type ShapeText,
  type SheetShape,
  type TextRun,
  type Theme,
} from "@cells/core"
import { shapeGeometry } from "./path"
import { presetGeometry } from "./geometry"
import { pictureFilter, pictureOpacity } from "./picture-filter"

/**
 * Shapes, text boxes, pictures and groups drawn as SVG markup at their size (pixels at 100%), for the sheet, for
 * printing and for copying as a picture. Text is laid out here (wrapped to the shape's text box), so the SVG has no
 * HTML in it and draws onto a canvas.
 */

/** Text width in pixels for a CSS font; the view passes a canvas's, tests an estimate. */
export type Measure = (text: string, font: string) => number

export const estimate: Measure = (text, font) => {
  const px = Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? 14)
  return text.length * px * 0.52
}

let canvas: CanvasRenderingContext2D | null | undefined
/** A canvas measure where there is one (the app), the estimate elsewhere. */
export const canvasMeasure: Measure = (text, font) => {
  if (canvas === undefined) {
    try {
      canvas = typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d")
    } catch {
      canvas = null
    }
  }
  if (!canvas) return estimate(text, font)
  canvas.font = font
  return canvas.measureText(text).width
}

import { slicerSvg, slicerView } from "./slicer"
import { timelineSvg, timelineView } from "./timeline"

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
const n = (v: number) => Math.round(v * 100) / 100
const PT = 96 / 72

function paint(c: ShapeColor | null | undefined, theme: Theme, attr: "fill" | "stroke"): string {
  if (!c) return ` ${attr}="none"`
  const { hex, alpha } = colorCss(c, theme)
  return ` ${attr}="${hex}"${alpha < 1 ? ` ${attr}-opacity="${n(alpha)}"` : ""}`
}

const fillAttr = (f: ShapeFill | undefined, theme: Theme) =>
  f?.kind === "solid" ? paint(f.color, theme, "fill") : ` fill="none"`

const DASHES: Record<string, (w: number) => string> = {
  dot: (w) => `${w} ${w}`,
  sysDot: (w) => `${w} ${w}`,
  dash: (w) => `${w * 4} ${w * 3}`,
  sysDash: (w) => `${w * 3} ${w}`,
  lgDash: (w) => `${w * 8} ${w * 3}`,
  dashDot: (w) => `${w * 4} ${w * 3} ${w} ${w * 3}`,
}

function strokeAttr(l: ShapeLine | undefined, theme: Theme): string {
  if (!l?.color || l.width <= 0) return ` stroke="none"`
  const w = Math.max(0.75, l.width * PT)
  const dash = l.dash && DASHES[l.dash] ? ` stroke-dasharray="${DASHES[l.dash](w)}"` : ""
  return `${paint(l.color, theme, "stroke")} stroke-width="${n(w)}" stroke-linejoin="round"${dash}`
}

/** Arrowheads for a line's ends, as SVG markers. */
function markers(id: string, l: ShapeLine, theme: Theme): { defs: string; attrs: string } {
  const color = l.color ? colorCss(l.color, theme).hex : "#000"
  let defs = ""
  let attrs = ""
  for (const [end, kind] of [
    ["start", l.head],
    ["end", l.tail],
  ] as const) {
    if (!kind || kind === "none") continue
    const mid = `${id}-${end}`
    const shape =
      kind === "oval"
        ? `<circle cx="5" cy="5" r="4"/>`
        : kind === "diamond"
          ? `<path d="M0 5 L5 0 L10 5 L5 10 Z"/>`
          : kind === "arrow"
            ? `<path d="M0 0 L10 5 L0 10" fill="none" stroke="${color}" stroke-width="1.5"/>`
            : kind === "stealth"
              ? `<path d="M0 0 L10 5 L0 10 L3 5 Z"/>`
              : `<path d="M0 0 L10 5 L0 10 Z"/>`
    defs += `<marker id="${mid}" viewBox="0 0 10 10" refX="${kind === "triangle" || kind === "stealth" ? 9 : 5}" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse" fill="${color}">${shape}</marker>`
    attrs += ` marker-${end}="url(#${mid})"`
  }
  return { defs, attrs }
}

// ── Text ─────────────────────────────────────────────────────────────────────────────────────────────────────

type Line = {
  runs: { run: TextRun; text: string; w: number }[]
  w: number
  h: number
  align: string
  /** An equation's line (Insert › Equation), drawn in the paragraph's colour. */
  math?: { box: MathBox; color?: ShapeColor }
}

const runFont = (r: TextRun, theme: Theme) =>
  `${r.italic ? "italic " : ""}${r.bold ? "700 " : ""}${n((r.size ?? 11) * PT)}px "${r.font ?? theme.fonts.minor}", "Aptos", "Calibri", sans-serif`

/** Paragraphs broken into lines of at most `width` pixels (no wrapping: one line each). */
export function layoutText(text: ShapeText, width: number, theme: Theme, measure: Measure): Line[] {
  const lines: Line[] = []
  for (const p of text.paragraphs) {
    if (p.math) {
      const look = p.runs[0]
      const px = (look?.size ?? 11) * PT
      if (p.math.every((l) => !l.length)) {
        // A new, empty equation: Office's prompt, in grey.
        const prompt: TextRun = {
          ...look,
          text: "Type equation here.",
          color: { theme: "dk1", lumMod: 0.5, lumOff: 0.5 },
        }
        const w = measure(prompt.text, runFont(prompt, theme))
        lines.push({ runs: [{ run: prompt, text: prompt.text, w }], w, h: px * 1.2, align: "ctr" })
        continue
      }
      for (const box of layoutEquation(p.math, px, "currentColor", measure))
        lines.push({
          runs: [],
          w: box.w,
          h: box.a + box.d + px * 0.3,
          align: p.align ?? "ctr",
          math: { box, color: look?.color },
        })
      continue
    }
    const align = p.align ?? "l"
    const size = Math.max(...p.runs.map((r) => r.size ?? 11), p.runs.length ? 0 : 11)
    const lineH = size * PT * 1.2
    let line: Line = { runs: [], w: 0, h: lineH, align }
    const push = () => {
      lines.push(line)
      line = { runs: [], w: 0, h: lineH, align }
    }
    for (const run of p.runs) {
      const font = runFont(run, theme)
      for (const word of run.text.split(/(?<=\s)/)) {
        const w = measure(word, font)
        if (text.wrap !== false && line.w + w > width && line.w > 0 && word.trim()) push()
        const last = line.runs[line.runs.length - 1]
        if (last && last.run === run) {
          last.text += word
          last.w += w
        } else line.runs.push({ run, text: word, w })
        line.w += w
      }
    }
    push()
  }
  return lines
}

/** The box an equation's text needs (its insets included), to fit the shape to it as Office does. */
export function textSize(text: ShapeText, theme: Theme, measure: Measure = canvasMeasure): { w: number; h: number } {
  const ins = text.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }
  const lines = layoutText({ ...text, wrap: false }, Infinity, theme, measure)
  const w = Math.max(24, ...lines.map((l) => l.w))
  const h = lines.reduce((sum, l) => sum + l.h, 0)
  return { w: Math.ceil(w + ins.l + ins.r + 2), h: Math.ceil(h + ins.t + ins.b) }
}

function textSvg(
  text: ShapeText,
  box: { x: number; y: number; w: number; h: number },
  theme: Theme,
  measure: Measure,
  wordArt?: SheetShape["wordArt"],
  defaultColor?: ShapeColor,
): string {
  const ins = text.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }
  const vertical = text.vertical === "vert" || text.vertical === "vert270"
  const inner = {
    x: box.x + ins.l,
    y: box.y + ins.t,
    w: Math.max(1, box.w - ins.l - ins.r),
    h: Math.max(1, box.h - ins.t - ins.b),
  }
  const lines = layoutText(text, vertical ? inner.h : inner.w, theme, measure)
  if (!lines.some((l) => l.math || l.runs.some((r) => r.text.trim()))) return ""
  const total = lines.reduce((s, l) => s + l.h, 0)
  const room = vertical ? inner.w : inner.h
  let y = text.anchor === "ctr" ? (room - total) / 2 : text.anchor === "b" ? room - total : 0
  const width = vertical ? inner.h : inner.w
  const out: string[] = []
  const effects = wordArt?.effects
  for (const line of lines) {
    const x0 = line.align === "ctr" ? (width - line.w) / 2 : line.align === "r" ? width - line.w : 0
    let x = x0
    if (line.math) {
      const { box, color } = line.math
      const hex = colorCss(color ?? defaultColor ?? { theme: "dk1" }, theme).hex
      out.push(
        `<g transform="translate(${n(x0)} ${n(y + (line.h - box.a - box.d) / 2 + box.a)})" color="${hex}">${box.svg}</g>`,
      )
      y += line.h
      continue
    }
    const base = y + line.h * 0.8
    const spans = line.runs.map((r) => {
      const color = r.run.color ?? defaultColor ?? { theme: "dk1" as const }
      const fill = wordArt?.fill ? fillAttr(wordArt.fill, theme) : paint(color, theme, "fill")
      const stroke = effects?.outline ? strokeAttr(effects.outline, theme) + ` paint-order="stroke"` : ""
      const deco = [r.run.underline ? "underline" : "", r.run.strike ? "line-through" : ""].filter(Boolean).join(" ")
      const span = `<tspan x="${n(x)}" y="${n(base)}"${fill}${stroke}${deco ? ` text-decoration="${deco}"` : ""} style="font:${esc(runFont(r.run, theme))}">${esc(r.text)}</tspan>`
      x += r.w
      return span
    })
    out.push(`<text>${spans.join("")}</text>`)
    y += line.h
  }
  const rotate =
    text.vertical === "vert"
      ? ` transform="translate(${n(inner.x + inner.w)} ${n(inner.y)}) rotate(90)"`
      : text.vertical === "vert270"
        ? ` transform="translate(${n(inner.x)} ${n(inner.y + inner.h)}) rotate(-90)"`
        : ` transform="translate(${n(inner.x)} ${n(inner.y)})"`
  const filter = effects?.shadow ? ` filter="url(#text-shadow)"` : ""
  return `<g${rotate}${filter} xml:space="preserve">${out.join("")}</g>`
}

// ── Objects ──────────────────────────────────────────────────────────────────────────────────────────────────

const SHADOW = `<filter id="shape-shadow" x="-20%" y="-20%" width="150%" height="150%"><feDropShadow dx="2" dy="2" stdDeviation="2" flood-opacity="0.4"/></filter><filter id="text-shadow" x="-20%" y="-20%" width="150%" height="150%"><feDropShadow dx="1.5" dy="1.5" stdDeviation="1" flood-opacity="0.4"/></filter>`

const looks = new Map<string, EffectLook[]>()

/** A shape's theme effect as an SVG filter (its shadow and glow), or none when the effect has neither. */
export function themeEffectFilter(id: string, effect: 1 | 2 | 3, theme: Theme, ownColor: string): string {
  const scheme = theme.effects?.scheme ?? ""
  let list = looks.get(scheme)
  if (!list) looks.set(scheme, (list = effectLooks(scheme || undefined)))
  const look = list[effect - 1]
  const parts: string[] = []
  const merge: string[] = []
  if (look.glow && look.glow.radius > 0) {
    const g = look.glow
    parts.push(
      `<feMorphology in="SourceAlpha" operator="dilate" radius="${n(g.radius / 2)}" result="gm"/>` +
        `<feGaussianBlur in="gm" stdDeviation="${n(g.radius / 2)}" result="gb"/>` +
        `<feFlood flood-color="${g.color ?? ownColor}" flood-opacity="${n(g.alpha)}"/>` +
        `<feComposite in2="gb" operator="in" result="glow"/>`,
    )
    merge.push("glow")
  }
  if (look.shadow && !look.shadow.inner) {
    const s = look.shadow
    parts.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${n(s.blur / 2)}"/><feOffset dx="${n(s.dx)}" dy="${n(s.dy)}" result="so"/>` +
        `<feFlood flood-color="${s.color}" flood-opacity="${n(s.alpha)}"/><feComposite in2="so" operator="in" result="shadow"/>`,
    )
    merge.push("shadow")
  }
  let inner = ""
  if (look.shadow?.inner) {
    const s = look.shadow
    inner =
      `<feComponentTransfer in="SourceAlpha"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer>` +
      `<feGaussianBlur stdDeviation="${n(s.blur / 2)}"/><feOffset dx="${n(s.dx)}" dy="${n(s.dy)}" result="io"/>` +
      `<feFlood flood-color="${s.color}" flood-opacity="${n(s.alpha)}"/><feComposite in2="io" operator="in"/>` +
      `<feComposite in2="SourceAlpha" operator="in" result="inner"/>`
  }
  if (!parts.length && !inner) return ""
  const nodes = [...merge, "SourceGraphic", ...(inner ? ["inner"] : [])].map((m) => `<feMergeNode in="${m}"/>`)
  return `<filter id="${id}" x="-30%" y="-30%" width="160%" height="160%">${parts.join("")}${inner}<feMerge>${nodes.join("")}</feMerge></filter>`
}

/** An object's SVG content in its own box (0,0)–(w,h), not rotated. */
export function objectSvg(o: DrawingObject, w: number, h: number, theme: Theme, measure: Measure = estimate): string {
  if (o.kind === "picture" && o.signature) return signatureSvg(o.signature, w, h, theme)
  if (o.kind === "picture" && o.embed) return embedSvg(o.embed, w, h, theme)
  if (o.kind === "picture") {
    const c = o.crop ?? { l: 0, t: 0, r: 0, b: 0 }
    const fw = w / Math.max(0.01, 1 - c.l - c.r)
    const fh = h / Math.max(0.01, 1 - c.t - c.b)
    const href = o.svg ?? o.src
    // Crop to Shape: the picture is cut to the shape, and its border follows it.
    const shape = o.geometry ? presetGeometry(o.geometry, w, h) : null
    const outline = shape?.paths.join(" ")
    const clipId = `pic${o.id}clip`
    // Color and Transparency (Office 2007's adjustments), drawn as a filter and the picture's opacity.
    const fxId = `pic${o.id}fx`
    const fx = pictureFilter(fxId, o.adjust, theme)
    const opacity = pictureOpacity(o.adjust)
    const defs = (outline ? `<clipPath id="${clipId}"><path d="${outline}" fill-rule="evenodd"/></clipPath>` : "") + fx
    const clip = defs ? `<defs>${defs}</defs>` : ""
    const look = (fx ? ` filter="url(#${fxId})"` : "") + (opacity !== null ? ` opacity="${n(opacity)}"` : "")
    const line = !o.line
      ? ""
      : outline
        ? `<path d="${outline}" fill="none" fill-rule="evenodd"${strokeAttr(o.line, theme)}/>`
        : `<rect x="0" y="0" width="${n(w)}" height="${n(h)}" fill="none"${strokeAttr(o.line, theme)}/>`
    return (
      `${clip}<svg x="0" y="0" width="${n(w)}" height="${n(h)}" overflow="hidden"><image href="${esc(href)}" x="${n(-c.l * fw)}" y="${n(-c.t * fh)}" width="${n(fw)}" height="${n(fh)}" preserveAspectRatio="none"${look}${outline ? ` clip-path="url(#${clipId})"` : ""}/></svg>` +
      line
    )
  }
  if (o.kind === "slicer") return slicerSvg(o, slicerView(o), w, h, theme, measure)
  if (o.kind === "timeline") return timelineSvg(o, timelineView(o), w, h, theme, measure)
  if (o.kind === "group") {
    const sx = w / (o.child.w || 1)
    const sy = h / (o.child.h || 1)
    const items = o.items
      .filter((i): i is DrawingObject => !isChart(i) && !i.hidden)
      .map((i) => {
        const b = i.box ?? { x: 0, y: 0, w: o.child.w, h: o.child.h }
        return placed(i, { x: b.x * sx, y: b.y * sy, w: b.w * sx, h: b.h * sy }, theme, measure)
      })
    return items.join("")
  }
  if (o.ink) return inkSvg(o.ink.strokes, w, h)
  const g = shapeGeometry(o, w, h)
  const id = `o${o.id}`
  const stroke = strokeAttr(o.line, theme)
  const ends = o.connector ? markers(id, o.line, theme) : { defs: "", attrs: "" }
  const fx =
    o.themeEffect && !o.shadow
      ? themeEffectFilter(
          `${id}fx`,
          o.themeEffect,
          theme,
          o.fill.kind === "solid" ? colorCss(o.fill.color, theme).hex : "#156082",
        )
      : ""
  const filter = o.shadow ? ` filter="url(#shape-shadow)"` : fx ? ` filter="url(#${id}fx)"` : ""
  const fill = g.open || o.connector ? ` fill="none"` : fillAttr(o.fill, theme)
  const shapes =
    g.paths.map((d) => `<path d="${d}" fill-rule="evenodd"${fill}${stroke}${ends.attrs}${filter}/>`).join("") +
    (g.lines ?? []).map((d) => `<path d="${d}" fill="none"${stroke}/>`).join("")
  const text = o.text && !o.connector ? textSvg(o.text, g.text, theme, measure, o.wordArt) : ""
  const defs = ends.defs + fx
  return `${defs ? `<defs>${defs}</defs>` : ""}${shapes}${text}`
}

/** A smooth path through points (x, y pairs): quadratic curves through the midpoints. */
export function inkPath(p: readonly number[]): string {
  if (p.length < 4) return `M${n(p[0])} ${n(p[1])}l0.01 0`
  let d = `M${n(p[0])} ${n(p[1])}`
  for (let i = 2; i + 3 < p.length; i += 2)
    d += `Q${n(p[i])} ${n(p[i + 1])} ${n((p[i] + p[i + 2]) / 2)} ${n((p[i + 1] + p[i + 3]) / 2)}`
  return d + `L${n(p[p.length - 2])} ${n(p[p.length - 1])}`
}

/** Ink strokes in a box (points are fractions of it): pens round, highlighters flat-tipped and see-through. */
export function inkSvg(strokes: readonly InkStroke[], w: number, h: number): string {
  return strokes
    .map((s) => {
      const d = inkPath(s.points.map((v, i) => v * (i % 2 ? h : w)))
      const look = s.highlighter
        ? ` stroke-linecap="square" stroke-linejoin="round" opacity="0.5" style="mix-blend-mode:multiply"`
        : ` stroke-linecap="round" stroke-linejoin="round"`
      return `<path d="${d}" fill="none" stroke="${esc(s.color)}" stroke-width="${n(s.width)}"${look}/>`
    })
    .join("")
}

/** An object at a box, turned and flipped about its middle. */
export function placed(
  o: DrawingObject,
  box: { x: number; y: number; w: number; h: number },
  theme: Theme,
  measure: Measure = estimate,
): string {
  const cx = box.w / 2
  const cy = box.h / 2
  const t = [`translate(${n(box.x)} ${n(box.y)})`]
  if (o.rotation) t.push(`rotate(${n(o.rotation)} ${n(cx)} ${n(cy)})`)
  if (o.flipH || o.flipV)
    t.push(
      `translate(${o.flipH ? n(box.w) : 0} ${o.flipV ? n(box.h) : 0}) scale(${o.flipH ? -1 : 1} ${o.flipV ? -1 : 1})`,
    )
  return `<g transform="${t.join(" ")}">${objectSvg(o, box.w, box.h, theme, measure)}</g>`
}

/** A whole SVG document for one object (its own size), for the sheet's layer, printing and copying. */
export function objectDocument(
  o: DrawingObject,
  w: number,
  h: number,
  theme: Theme,
  measure: Measure = estimate,
): string {
  return `<defs>${SHADOW}</defs>${objectSvg(o, w, h, theme, measure)}`
}

/**
 * An object as a group in a larger SVG, at (x, y) with its rotation and flips: for copying several objects as one
 * picture and for printing.
 */
export function objectAt(
  o: DrawingObject,
  box: { x: number; y: number; w: number; h: number },
  theme: Theme,
  measure: Measure = estimate,
): string {
  const cx = box.w / 2
  const cy = box.h / 2
  let t = `translate(${box.x} ${box.y})`
  if (o.rotation) t += ` rotate(${o.rotation} ${cx} ${cy})`
  if (o.flipH || o.flipV)
    t += ` translate(${cx} ${cy}) scale(${o.flipH ? -1 : 1} ${o.flipV ? -1 : 1}) translate(${-cx} ${-cy})`
  return `<g transform="${t}">${objectDocument(o, box.w, box.h, theme, measure)}</g>`
}
