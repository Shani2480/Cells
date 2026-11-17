import {
  drawnLayout,
  smartArtColors,
  type SheetShape,
  type ShapeColor,
  type ShapeText,
  type SmartArt,
  type SmartArtNode,
  type Theme,
} from "@cells/core"
import { canvasMeasure, layoutText, type Measure } from "./render"

/**
 * Insert › SmartArt: Cells' layouts laid out as Excel lays out their definitions (core `smartart.ts`), measured
 * against Excel: the same places and sizes for any number of points, one text size for all of them (the largest
 * from 65 pt down to 5 pt at which every point's text fits, words unbroken).
 */

type Box = { x: number; y: number; w: number; h: number }
const PT = 96 / 72
/** Office's text sizes: the most a layout starts with, the least it shrinks to. */
const MAX_FONT = 65
const MIN_FONT = 5

/** Where a layout puts its shapes, for points' boxes and the arrows between them (`rot`: degrees). */
type Placed = { nodes: Box[]; arrows: (Box & { rot: number })[]; geometry: string; adjust?: Record<string, number> }

/** Blocks in rows (Basic Block List, Grid Matrix): `ratio` height to width, `gap` × width apart, `cols` fixed. */
function snake(n: number, size: { w: number; h: number }, ratio: number, gap: number, cols?: number): Box[] {
  let best = { w: 0, cols: 1 }
  for (let c = cols ?? 1; c <= (cols ?? n); c++) {
    const rows = Math.ceil(n / c)
    const w = Math.min(size.w / (c + (c - 1) * gap), size.h / (rows * ratio + (rows - 1) * gap))
    if (w > best.w + 1e-9) best = { w, cols: c }
  }
  const { w, cols: c } = best
  const h = w * ratio
  const rows = Math.ceil(n / c)
  const height = rows * h + (rows - 1) * gap * w
  const y0 = (size.h - height) / 2
  const out: Box[] = []
  for (let i = 0; i < n; i++) {
    const row = Math.floor(i / c)
    // A short last row is centred.
    const inRow = Math.min(c, n - row * c)
    const width = inRow * w + (inRow - 1) * gap * w
    const x0 = (size.w - width) / 2
    out.push({ x: x0 + (i % c) * w * (1 + gap), y: y0 + row * (h + gap * w), w, h })
  }
  return out
}

/** Steps in a row with an arrow in each gap (Basic Process). */
function process(n: number, size: { w: number; h: number }): Placed {
  const slot = 0.35
  const w = Math.min(size.w / (n + (n - 1) * slot), size.h / 0.6)
  const h = w * 0.6
  const width = n * w + (n - 1) * slot * w
  const x0 = (size.w - width) / 2
  const y = (size.h - h) / 2
  const nodes = Array.from({ length: n }, (_, i) => ({ x: x0 + i * w * (1 + slot), y, w, h }))
  const gap = slot * w
  const arrows = nodes.slice(0, -1).map((b) => {
    const aw = gap * (1 - 0.25 - 0.22)
    const ah = gap * 0.62
    return { x: b.x + b.w + gap * 0.25, y: b.y + h / 2 - ah / 2, w: aw, h: ah, rot: 0 }
  })
  return { nodes, arrows, geometry: "roundRect", adjust: { adj: 10000 } }
}

/** Circles on a ring from the top, clockwise, an arrow along each chord, the last back to the first (Basic Cycle). */
function cycle(n: number, size: { w: number; h: number }): Placed {
  const sp = 0.35
  // For a circle 1 across: the ring's radius, then the whole drawing's extent, scaled to fit.
  const r = n > 1 ? (1 + sp) / (2 * Math.sin(Math.PI / n)) : 0
  const at = (i: number) => {
    const a = (2 * Math.PI * i) / n
    return { x: r * Math.sin(a), y: -r * Math.cos(a) }
  }
  const centres = Array.from({ length: n }, (_, i) => at(i))
  const x0 = Math.min(...centres.map((c) => c.x)) - 0.5
  const x1 = Math.max(...centres.map((c) => c.x)) + 0.5
  const y0 = Math.min(...centres.map((c) => c.y)) - 0.5
  const y1 = Math.max(...centres.map((c) => c.y)) + 0.5
  const d = Math.min(size.w / (x1 - x0), size.h / (y1 - y0))
  const ox = (size.w - (x1 - x0) * d) / 2 - x0 * d
  const oy = (size.h - (y1 - y0) * d) / 2 - y0 * d
  const nodes = centres.map((c) => ({ x: ox + c.x * d - d / 2, y: oy + c.y * d - d / 2, w: d, h: d }))
  const arrows =
    n > 1
      ? centres.map((a, i) => {
          const b = centres[(i + 1) % n]!
          const dx = (b.x - a.x) * d
          const dy = (b.y - a.y) * d
          const gap = Math.hypot(dx, dy) - d
          const aw = gap * 0.7
          const ah = 0.24 * d
          const cx = ox + ((a.x + b.x) / 2) * d
          const cy = oy + ((a.y + b.y) / 2) * d
          const rot = ((Math.atan2(dy, dx) * 180) / Math.PI + 360) % 360
          return { x: cx - aw / 2, y: cy - ah / 2, w: aw, h: ah, rot: Math.round(rot * 100) / 100 }
        })
      : []
  return { nodes, arrows, geometry: "ellipse" }
}

function place(key: string, n: number, size: { w: number; h: number }): Placed {
  if (key === "process") return process(n, size)
  if (key === "cycle") return cycle(n, size)
  if (key === "matrix") return { nodes: snake(n, size, 1, 0.05, Math.min(2, n)), arrows: [], geometry: "rect" }
  return { nodes: snake(n, size, 0.6, 0.1), arrows: [], geometry: "rect" }
}

/** A line of a point's shape: its own text (level 0), then the points under it (1, 2…). */
type Line = { text: string; level: number }

/** A point's text as its shape shows it: its own lines, then the points under it as bullets. */
function lines(node: SmartArtNode): Line[] {
  const out: Line[] = node.text ? node.text.split("\n").map((text) => ({ text, level: 0 })) : []
  const under = (list: SmartArtNode[], level: number) => {
    for (const c of list) {
      if (c.text) out.push({ text: c.text, level })
      under(c.children, level + 1)
    }
  }
  under(node.children, 1)
  return out
}

/** Office's size for the points under a point, against the point's own. */
const SUB_SIZE = 0.78

/**
 * Text in a shape at `size` points, with Office's margins (0.3 × the size on each side): centred, or with points
 * under it at the top left, those as smaller bullets, as Office lays them out.
 */
function nodeText(text: Line[], size: number, color: ShapeColor): ShapeText {
  const inset = 0.3 * size * PT
  const bullets = text.some((l) => l.level > 0)
  return {
    paragraphs: (text.length ? text : [{ text: "", level: 0 }]).map((l) => {
      const sub = Math.round(size * SUB_SIZE * 2) / 2
      const run = {
        text: l.level ? `${"\u2003".repeat(l.level - 1)}\u2022 ${l.text}` : l.text,
        size: l.level ? sub : size,
        color,
      }
      return { align: bullets ? "l" : "ctr", runs: [run] }
    }),
    anchor: bullets ? "t" : "ctr",
    wrap: true,
    insets: { l: inset, t: inset, r: inset, b: inset },
  }
}

/** Whether every point's text fits its box at `size` points, no word broken. */
function fits(texts: Line[][], boxes: Box[], size: number, theme: Theme, measure: Measure): boolean {
  return texts.every((t, i) => {
    const b = boxes[i]!
    if (!t.length) return true
    const text = nodeText(t, size, {})
    const inner = { w: b.w - 2 * 0.3 * size * PT, h: b.h - 2 * 0.3 * size * PT }
    if (inner.w <= 0 || inner.h <= 0) return false
    for (const p of text.paragraphs) {
      const r = p.runs[0]!
      const font = `${r.size! * PT}px "${theme.fonts.minor}", "Aptos", "Calibri", sans-serif`
      if (r.text.split(/\s+/).some((word) => measure(word, font) > inner.w)) return false
    }
    const laid = layoutText(text, inner.w, theme, measure)
    return laid.reduce((sum, l) => sum + l.h, 0) <= inner.h
  })
}

/**
 * A SmartArt's shapes laid out in a frame of `size` pixels: a shape for each top-level point (the points under it
 * as bullets in it) and the layout's arrows, in the SmartArt's colours. Empty points show Office's "[Text]".
 */
export function layoutSmartArt(
  sa: SmartArt,
  size: { w: number; h: number },
  theme: Theme,
  measure: Measure = canvasMeasure,
): SheetShape[] {
  const layout = drawnLayout(sa.layout)
  const n = Math.max(1, sa.nodes.length)
  const placed = place(layout.key, n, size)
  // Empty points show Office's place to type.
  const texts = sa.nodes.map((node) => {
    const t = lines(node)
    return t.length ? t : null
  })
  const blank: Line[] = [{ text: "[Text]", level: 0 }]
  const shown = texts.map((t) => t ?? blank)
  // One size for every point: the largest whose text fits, in Office's steps.
  let font = MAX_FONT
  while (font > MIN_FONT && !fits(shown, placed.nodes, font, theme, measure)) font -= font > 20 ? 1 : 0.5
  const colors = smartArtColors(sa.colors)
  const accent = (i: number) => colors.accents[i % colors.accents.length]! as ShapeColor["theme"]
  const shapes: SheetShape[] = []
  placed.nodes.forEach((box, i) => {
    const t = texts[i]
    shapes.push({
      kind: "shape",
      id: i + 1,
      name: "",
      anchor: { from: { row: 0, col: 0, x: 0, y: 0 }, to: { row: 0, col: 0, x: 0, y: 0 } },
      box,
      geometry: placed.geometry,
      ...(placed.adjust ? { adjust: placed.adjust } : {}),
      fill: { kind: "solid", color: { theme: accent(i) } },
      line: { color: { theme: "lt1" }, width: 1.5 },
      // An empty point: Office's place to type, in the text's colour.
      text: nodeText(t ?? blank, font, t ? { theme: "lt1" } : { theme: "lt1", alpha: 0.6 }),
    })
  })
  placed.arrows.forEach((box, i) => {
    const { rot, ...b } = box
    shapes.push({
      kind: "shape",
      id: placed.nodes.length + i + 1,
      name: "",
      anchor: { from: { row: 0, col: 0, x: 0, y: 0 }, to: { row: 0, col: 0, x: 0, y: 0 } },
      box: b,
      ...(rot ? { rotation: rot } : {}),
      geometry: "rightArrow",
      adjust: { adj1: 60000, adj2: 50000 },
      fill: { kind: "solid", color: { theme: accent(i), lumMod: 0.4, lumOff: 0.6 } },
      line: { color: null, width: 0 },
    })
  })
  return shapes
}
