/**
 * Excel's preset shapes as SVG paths, at a size: the shapes in Insert › Shapes, with their adjust values (the yellow
 * handles) as DrawingML defines them (100000ths of the shorter side unless noted). A shape Cells doesn't know draws
 * as a rectangle; its file's own geometry is kept.
 */

export type Geometry = {
  /** One or more SVG path strings (a can's top is drawn apart, so it reads as a lid). */
  paths: string[]
  /** Paths drawn without fill (a can's rim, a smiley's mouth). */
  lines?: string[]
  /** Where text goes, inside the shape. */
  text: { x: number; y: number; w: number; h: number }
  /** A line or arrow: no fill. */
  open?: boolean
}

type Adj = Record<string, number>

const f = (n: number) => Math.round(n * 100) / 100
const poly = (pts: [number, number][]) => "M" + pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(" L") + " Z"

function ellipseArc(cx: number, cy: number, rx: number, ry: number) {
  return `M${f(cx - rx)} ${f(cy)} A${f(rx)} ${f(ry)} 0 1 0 ${f(cx + rx)} ${f(cy)} A${f(rx)} ${f(ry)} 0 1 0 ${f(cx - rx)} ${f(cy)} Z`
}

function star(w: number, h: number, points: number, inner: number) {
  const pts: [number, number][] = []
  for (let i = 0; i < points * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / points
    const r = i % 2 ? inner : 0.5
    pts.push([w / 2 + Math.cos(a) * r * w, h / 2 + Math.sin(a) * r * h])
  }
  return poly(pts)
}

function regular(w: number, h: number, sides: number, rotate = 0) {
  const pts: [number, number][] = []
  for (let i = 0; i < sides; i++) {
    const a = -Math.PI / 2 + rotate + (i * 2 * Math.PI) / sides
    pts.push([w / 2 + (Math.cos(a) * w) / 2, h / 2 + (Math.sin(a) * h) / 2])
  }
  return poly(pts)
}

/** The geometry of a preset at w × h pixels. */
export function presetGeometry(name: string, w: number, h: number, adjust: Adj = {}): Geometry {
  const ss = Math.min(w, h)
  const a = (key: string, d: number) => (adjust[key] ?? d) / 100000
  const full = { x: 0, y: 0, w, h }
  switch (name) {
    case "line":
    case "straightConnector1":
    case "bentConnector3":
    case "curvedConnector3":
      return { paths: [`M0 0 L${f(w)} ${f(h)}`], text: full, open: true }
    case "ellipse":
    case "flowChartConnector":
      return { paths: [ellipseArc(w / 2, h / 2, w / 2, h / 2)], text: inset(w, h, 0.146) }
    case "roundRect":
    case "flowChartAlternateProcess": {
      const r = Math.min(ss * a("adj", 16667), ss / 2)
      return {
        paths: [
          `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)} V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`,
        ],
        text: { x: r * 0.29, y: r * 0.29, w: w - r * 0.58, h: h - r * 0.58 },
      }
    }
    case "snip1Rect": {
      const s = ss * a("adj", 16667)
      return {
        paths: [
          poly([
            [0, 0],
            [w - s, 0],
            [w, s],
            [w, h],
            [0, h],
          ]),
        ],
        text: full,
      }
    }
    case "triangle": {
      const x = w * a("adj", 50000)
      return {
        paths: [
          poly([
            [x, 0],
            [w, h],
            [0, h],
          ]),
        ],
        text: { x: x / 2, y: h / 2, w: w / 2, h: h / 2 },
      }
    }
    case "rtTriangle":
      return {
        paths: [
          poly([
            [0, 0],
            [w, h],
            [0, h],
          ]),
        ],
        text: { x: w / 12, y: h * 0.58, w: w / 2, h: h * 0.42 },
      }
    case "parallelogram": {
      const x = ss * a("adj", 25000)
      return {
        paths: [
          poly([
            [x, 0],
            [w, 0],
            [w - x, h],
            [0, h],
          ]),
        ],
        text: { x: x, y: 0, w: w - 2 * x, h },
      }
    }
    case "trapezoid": {
      const x = ss * a("adj", 25000)
      return {
        paths: [
          poly([
            [x, 0],
            [w - x, 0],
            [w, h],
            [0, h],
          ]),
        ],
        text: { x: x / 2, y: h / 3, w: w - x, h: (h * 2) / 3 },
      }
    }
    case "diamond":
    case "flowChartDecision":
      return {
        paths: [
          poly([
            [w / 2, 0],
            [w, h / 2],
            [w / 2, h],
            [0, h / 2],
          ]),
        ],
        text: { x: w / 4, y: h / 4, w: w / 2, h: h / 2 },
      }
    case "pentagon":
      return { paths: [regular(w, h, 5)], text: inset(w, h, 0.2) }
    case "hexagon": {
      const x = ss * a("adj", 25000)
      return {
        paths: [
          poly([
            [x, 0],
            [w - x, 0],
            [w, h / 2],
            [w - x, h],
            [x, h],
            [0, h / 2],
          ]),
        ],
        text: { x, y: 0, w: w - 2 * x, h },
      }
    }
    case "octagon": {
      const x = ss * a("adj", 29289)
      return {
        paths: [
          poly([
            [x, 0],
            [w - x, 0],
            [w, x],
            [w, h - x],
            [w - x, h],
            [x, h],
            [0, h - x],
            [0, x],
          ]),
        ],
        text: { x: x / 2, y: x / 2, w: w - x, h: h - x },
      }
    }
    case "plus": {
      const x = ss * a("adj", 25000)
      return {
        paths: [
          poly([
            [x, 0],
            [w - x, 0],
            [w - x, x],
            [w, x],
            [w, h - x],
            [w - x, h - x],
            [w - x, h],
            [x, h],
            [x, h - x],
            [0, h - x],
            [0, x],
            [x, x],
          ]),
        ],
        text: { x: 0, y: x, w, h: h - 2 * x },
      }
    }
    case "star4":
      return { paths: [star(w, h, 4, 0.19)], text: inset(w, h, 0.3) }
    case "star5":
      return { paths: [star(w, h, 5, 0.19)], text: inset(w, h, 0.3) }
    case "star6":
      return { paths: [star(w, h, 6, 0.29)], text: inset(w, h, 0.25) }
    case "rightArrow":
    case "leftArrow": {
      const shaft = h * a("adj1", 50000)
      const head = Math.min(ss * a("adj2", 50000), w)
      const y0 = (h - shaft) / 2
      const pts: [number, number][] = [
        [0, y0],
        [w - head, y0],
        [w - head, 0],
        [w, h / 2],
        [w - head, h],
        [w - head, y0 + shaft],
        [0, y0 + shaft],
      ]
      return {
        paths: [poly(name === "leftArrow" ? pts.map(([x, y]) => [w - x, y]) : pts)],
        text: { x: 0, y: y0, w: w - head / 2, h: shaft },
      }
    }
    case "upArrow":
    case "downArrow": {
      const shaft = w * a("adj1", 50000)
      const head = Math.min(ss * a("adj2", 50000), h)
      const x0 = (w - shaft) / 2
      const pts: [number, number][] = [
        [x0, h],
        [x0, head],
        [0, head],
        [w / 2, 0],
        [w, head],
        [x0 + shaft, head],
        [x0 + shaft, h],
      ]
      return {
        paths: [poly(name === "downArrow" ? pts.map(([x, y]) => [x, h - y]) : pts)],
        text: { x: x0, y: head / 2, w: shaft, h: h - head / 2 },
      }
    }
    case "leftRightArrow": {
      const shaft = h * a("adj1", 50000)
      const head = Math.min(ss * a("adj2", 50000), w / 2)
      const y0 = (h - shaft) / 2
      return {
        paths: [
          poly([
            [0, h / 2],
            [head, 0],
            [head, y0],
            [w - head, y0],
            [w - head, 0],
            [w, h / 2],
            [w - head, h],
            [w - head, y0 + shaft],
            [head, y0 + shaft],
            [head, h],
          ]),
        ],
        text: { x: head / 2, y: y0, w: w - head, h: shaft },
      }
    }
    case "chevron": {
      const x = Math.min(ss * a("adj", 50000), w)
      return {
        paths: [
          poly([
            [0, 0],
            [w - x, 0],
            [w, h / 2],
            [w - x, h],
            [0, h],
            [x, h / 2],
          ]),
        ],
        text: { x, y: 0, w: w - 2 * x, h },
      }
    }
    case "homePlate": {
      const x = Math.min(ss * a("adj", 50000), w)
      return {
        paths: [
          poly([
            [0, 0],
            [w - x, 0],
            [w, h / 2],
            [w - x, h],
            [0, h],
          ]),
        ],
        text: { x: 0, y: 0, w: w - x / 2, h },
      }
    }
    case "heart":
      return {
        paths: [
          `M${f(w / 2)} ${f(h / 4)} C${f(w / 2)} ${f(-h / 12)} ${f(-w / 4)} ${f(h / 12)} ${f(w / 12)} ${f(h / 2)} L${f(w / 2)} ${f(h)} L${f((w * 11) / 12)} ${f(h / 2)} C${f((w * 5) / 4)} ${f(h / 12)} ${f(w / 2)} ${f(-h / 12)} ${f(w / 2)} ${f(h / 4)} Z`,
        ],
        text: inset(w, h, 0.2),
      }
    case "can": {
      const ry = Math.min(ss * a("adj", 25000), h / 2) / 2
      return {
        paths: [
          `M0 ${f(ry)} A${f(w / 2)} ${f(ry)} 0 0 1 ${f(w)} ${f(ry)} V${f(h - ry)} A${f(w / 2)} ${f(ry)} 0 0 1 0 ${f(h - ry)} Z`,
          ellipseArc(w / 2, ry, w / 2, ry),
        ],
        text: { x: 0, y: ry * 2, w, h: h - ry * 3 },
      }
    }
    case "donut": {
      const t = ss * a("adj", 25000)
      return {
        paths: [
          ellipseArc(w / 2, h / 2, w / 2, h / 2) +
            " " +
            ellipseArc(w / 2, h / 2, Math.max(0, w / 2 - t), Math.max(0, h / 2 - t)),
        ],
        text: inset(w, h, 0.146),
      }
    }
    case "frame": {
      const t = ss * a("adj1", 12500)
      return {
        paths: [`M0 0 H${f(w)} V${f(h)} H0 Z M${f(t)} ${f(t)} V${f(h - t)} H${f(w - t)} V${f(t)} Z`],
        text: full,
      }
    }
    case "wedgeRectCallout":
    case "wedgeRoundRectCallout":
    case "wedgeEllipseCallout": {
      const tx = w / 2 + w * a("adj1", -20833)
      const ty = h / 2 + h * a("adj2", 62500)
      if (name === "wedgeEllipseCallout") {
        const ang = Math.atan2(ty - h / 2, tx - w / 2)
        const p = (d: number) => [w / 2 + (Math.cos(ang + d) * w) / 2, h / 2 + (Math.sin(ang + d) * h) / 2]
        const [x1, y1] = p(0.18)
        const [x2, y2] = p(-0.18)
        return {
          paths: [`M${f(x1)} ${f(y1)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(x2)} ${f(y2)} L${f(tx)} ${f(ty)} Z`],
          text: inset(w, h, 0.146),
        }
      }
      // The tail leaves the side nearest its tip (the bottom by default).
      const bx = Math.min(Math.max(tx, w * 0.2), w * 0.8)
      const pts: [number, number][] =
        ty > h
          ? [
              [0, 0],
              [w, 0],
              [w, h],
              [bx + w * 0.08, h],
              [tx, ty],
              [bx - w * 0.08, h],
              [0, h],
            ]
          : ty < 0
            ? [
                [0, 0],
                [bx - w * 0.08, 0],
                [tx, ty],
                [bx + w * 0.08, 0],
                [w, 0],
                [w, h],
                [0, h],
              ]
            : [
                [0, 0],
                [w, 0],
                [w, h],
                [0, h],
              ]
      return { paths: [poly(pts)], text: full }
    }
    case "cloud":
    case "cloudCallout": {
      const bumps = [
        [0.25, 0.3, 0.22, 0.22],
        [0.5, 0.2, 0.25, 0.2],
        [0.75, 0.3, 0.22, 0.22],
        [0.8, 0.6, 0.2, 0.22],
        [0.55, 0.78, 0.25, 0.2],
        [0.28, 0.72, 0.22, 0.22],
        [0.18, 0.5, 0.17, 0.2],
      ]
      return {
        paths: [bumps.map(([x, y, rx, ry]) => ellipseArc(x * w, y * h, rx * w, ry * h)).join(" ")],
        text: inset(w, h, 0.2),
      }
    }
    case "flowChartTerminator": {
      const r = Math.min(h / 2, w / 2)
      return {
        paths: [
          `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(h / 2)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(h / 2)} 0 0 1 ${f(r)} 0 Z`,
        ],
        text: { x: r / 3, y: 0, w: w - (r * 2) / 3, h },
      }
    }
    case "smileyFace": {
      const e = Math.min(w, h) * 0.06
      return {
        paths: [ellipseArc(w / 2, h / 2, w / 2, h / 2)],
        lines: [
          ellipseArc(w * 0.35, h * 0.38, e, e),
          ellipseArc(w * 0.65, h * 0.38, e, e),
          `M${f(w * 0.28)} ${f(h * 0.65)} Q${f(w / 2)} ${f(h * 0.85)} ${f(w * 0.72)} ${f(h * 0.65)}`,
        ],
        text: inset(w, h, 0.146),
      }
    }
    case "lightningBolt":
      return {
        paths: [
          poly(
            [
              [0.39, 0],
              [0.62, 0.3],
              [0.53, 0.35],
              [0.8, 0.6],
              [0.7, 0.65],
              [1, 1],
              [0.48, 0.71],
              [0.58, 0.66],
              [0.22, 0.42],
              [0.36, 0.36],
              [0, 0.18],
            ].map(([x, y]) => [x * w, y * h]),
          ),
        ],
        text: inset(w, h, 0.3),
      }
    case "sun": {
      const rays: string[] = []
      for (let i = 0; i < 8; i++) {
        const ang = (i * Math.PI) / 4
        const c = (r: number, d = 0) => [w / 2 + Math.cos(ang + d) * r * w, h / 2 + Math.sin(ang + d) * r * h]
        rays.push(poly([c(0.5), c(0.34, 0.12), c(0.34, -0.12)] as [number, number][]))
      }
      return { paths: [ellipseArc(w / 2, h / 2, w * 0.25, h * 0.25), ...rays], text: inset(w, h, 0.3) }
    }
    case "moon": {
      const t = w * a("adj", 50000)
      return {
        paths: [`M${f(w)} 0 A${f(w)} ${f(h / 2)} 0 0 0 ${f(w)} ${f(h)} A${f(t)} ${f(h / 2)} 0 0 1 ${f(w)} 0 Z`],
        text: { x: t / 3, y: h / 4, w: t, h: h / 2 },
      }
    }
    default:
      // rect, flowChartProcess, text boxes, and shapes Cells doesn't know.
      return { paths: [`M0 0 H${f(w)} V${f(h)} H0 Z`], text: full }
  }
}

/**
 * A shape's adjustment handle (the yellow dot): where it is, and the adjust values a drag to a point gives (both in
 * the shape's own pixels, unturned), kept in the range Office allows.
 */
export type AdjustHandle = { x: number; y: number; drag: (x: number, y: number) => Adj }

/** The handles of a preset at w × h pixels with its adjust values (none for shapes without any). */
export function adjustHandles(name: string, w: number, h: number, adjust: Adj = {}): AdjustHandle[] {
  const ss = Math.max(1, Math.min(w, h))
  const a = (key: string, d: number) => (adjust[key] ?? d) / 100000
  const clamp = (v: number, max: number, min = 0) => Math.round(Math.min(max, Math.max(min, v)))
  /** One adjust value along the top edge: x pixels as a share of the shorter side. */
  const top = (key: string, x: number, max: number, from: (px: number) => number): AdjustHandle => ({
    x,
    y: 0,
    drag: (px) => ({ [key]: clamp((from(px) / ss) * 100000, max) }),
  })
  const wide = (k: number) => (k * 100000 * w) / ss
  switch (name) {
    case "roundRect":
    case "flowChartAlternateProcess": {
      const r = Math.min(ss * a("adj", 16667), ss / 2)
      return [top("adj", r, 50000, (px) => px)]
    }
    case "snip1Rect":
      return [top("adj", w - ss * a("adj", 16667), 50000, (px) => w - px)]
    case "triangle":
      return [{ x: w * a("adj", 50000), y: 0, drag: (px) => ({ adj: clamp((px / w) * 100000, 100000) }) }]
    case "parallelogram":
      return [top("adj", ss * a("adj", 25000), wide(1), (px) => px)]
    case "trapezoid":
    case "hexagon":
      return [top("adj", ss * a("adj", 25000), wide(0.5), (px) => px)]
    case "octagon":
      return [top("adj", ss * a("adj", 29289), 50000, (px) => px)]
    case "plus":
      return [top("adj", ss * a("adj", 25000), 50000, (px) => px)]
    case "frame":
      return [top("adj1", ss * a("adj1", 12500), 50000, (px) => px)]
    case "chevron":
    case "homePlate":
      return [top("adj", w - Math.min(ss * a("adj", 50000), w), wide(1), (px) => w - px)]
    case "donut": {
      const t = ss * a("adj", 25000)
      return [{ x: t, y: h / 2, drag: (px) => ({ adj: clamp((px / ss) * 100000, 50000) }) }]
    }
    case "can": {
      const y = Math.min(ss * a("adj", 25000), h / 2)
      return [{ x: w / 2, y, drag: (_px, py) => ({ adj: clamp((py / ss) * 100000, (50000 * h) / ss) }) }]
    }
    case "moon": {
      const t = w * a("adj", 50000)
      return [{ x: w - t, y: h / 2, drag: (px) => ({ adj: clamp(((w - px) / w) * 100000, 87500) }) }]
    }
    case "rightArrow":
    case "leftArrow":
    case "leftRightArrow": {
      const shaft = h * a("adj1", 50000)
      const head = Math.min(ss * a("adj2", 50000), name === "leftRightArrow" ? w / 2 : w)
      const left = name !== "rightArrow"
      const hx = left ? head : w - head
      const maxHead = name === "leftRightArrow" ? wide(0.5) : wide(1)
      return [
        { x: hx, y: (h - shaft) / 2, drag: (_px, py) => ({ adj1: clamp(((h - 2 * py) / h) * 100000, 100000) }) },
        { x: hx, y: 0, drag: (px) => ({ adj2: clamp(((left ? px : w - px) / ss) * 100000, maxHead) }) },
      ]
    }
    case "upArrow":
    case "downArrow": {
      const shaft = w * a("adj1", 50000)
      const head = Math.min(ss * a("adj2", 50000), h)
      const down = name === "downArrow"
      const hy = down ? h - head : head
      return [
        { x: (w - shaft) / 2, y: hy, drag: (px) => ({ adj1: clamp(((w - 2 * px) / w) * 100000, 100000) }) },
        {
          x: 0,
          y: hy,
          drag: (_px, py) => ({ adj2: clamp(((down ? h - py : py) / ss) * 100000, (100000 * h) / ss) }),
        },
      ]
    }
    case "wedgeRectCallout":
    case "wedgeRoundRectCallout":
    case "wedgeEllipseCallout":
      // The callout's point, anywhere.
      return [
        {
          x: w / 2 + w * a("adj1", -20833),
          y: h / 2 + h * a("adj2", 62500),
          drag: (px, py) => ({
            adj1: Math.round(((px - w / 2) / Math.max(1, w)) * 100000),
            adj2: Math.round(((py - h / 2) / Math.max(1, h)) * 100000),
          }),
        },
      ]
    default:
      return []
  }
}

/**
 * A preset's connection sites (where Excel glues connectors), in Office's order, at w × h pixels with its adjust
 * values. Measured in Excel 365 (each site of each shape at 200 × 100); shapes Cells draws its own way (the regular
 * pentagon, the stars) put theirs on Cells' corners.
 */
export function connectionSites(name: string, w: number, h: number, adjust: Adj = {}): [number, number][] {
  const ss = Math.min(w, h)
  const a = (key: string, d: number) => (adjust[key] ?? d) / 100000
  const T: [number, number] = [w / 2, 0]
  const L: [number, number] = [0, h / 2]
  const B: [number, number] = [w / 2, h]
  const R: [number, number] = [w, h / 2]
  const k = 0.1464
  const oval: [number, number][] = [
    T,
    [w * k, h * k],
    L,
    [w * k, h * (1 - k)],
    B,
    [w * (1 - k), h * (1 - k)],
    R,
    [w * (1 - k), h * k],
  ]
  const tip = (): [number, number] => [w / 2 + w * a("adj1", -20833), h / 2 + h * a("adj2", 62500)]
  /** The outer corners of Cells' regular shapes and stars, top first, then round the left side. */
  const corners = (points: number): [number, number][] => {
    const out: [number, number][] = []
    for (let i = 0; i < points; i++) {
      const t = -Math.PI / 2 - (i * 2 * Math.PI) / points
      out.push([w / 2 + (Math.cos(t) * w) / 2, h / 2 + (Math.sin(t) * h) / 2])
    }
    return out
  }
  switch (name) {
    case "line":
    case "straightConnector1":
    case "bentConnector3":
    case "curvedConnector3":
      return []
    case "snip1Rect":
      return [R, B, L, T]
    case "ellipse":
    case "flowChartConnector":
    case "donut":
    case "smileyFace":
      return oval
    case "wedgeEllipseCallout":
      return [...oval, tip()]
    case "wedgeRectCallout":
    case "wedgeRoundRectCallout":
      return [T, L, B, R, tip()]
    case "triangle": {
      const x = w * a("adj", 50000)
      return [
        [x, 0],
        [x / 2, h / 2],
        [0, h],
        [x, h],
        [w, h],
        [(x + w) / 2, h / 2],
      ]
    }
    case "rtTriangle":
      return [
        [0, 0],
        [0, h / 2],
        [0, h],
        [w / 2, h],
        [w, h],
        [w / 2, h / 2],
      ]
    case "parallelogram": {
      const x = ss * a("adj", 25000)
      return [T, [w / 2 + x / 2, 0], [w - x / 2, h / 2], [w / 2 - x / 2, h], B, [x / 2, h / 2]]
    }
    case "trapezoid": {
      const x = ss * a("adj", 25000)
      return [T, [x / 2, h / 2], B, [w - x / 2, h / 2]]
    }
    case "pentagon": {
      const [top, left, bl, br, right] = corners(5)
      return [top, left, bl, [w / 2, bl[1]], br, right]
    }
    case "hexagon": {
      const x = ss * a("adj", 25000)
      return [R, [w - x, h], [x, h], L, [x, 0], [w - x, 0]]
    }
    case "octagon": {
      const x = ss * a("adj", 29289)
      return [
        [w, x],
        [w, h - x],
        [w - x, h],
        [x, h],
        [0, h - x],
        [0, x],
        [x, 0],
        [w - x, 0],
      ]
    }
    case "can":
      return [[w / 2, Math.min(ss * a("adj", 25000), h / 2)], T, L, B, R]
    case "heart":
      return [[w / 2, h / 4], B]
    case "lightningBolt":
      return [
        [w * 0.3922, 0],
        [0, h * 0.1801],
        [w * 0.2325, h * 0.4493],
        [w * 0.4635, h * 0.6905],
        [w, h],
        [w * 0.7675, h * 0.5559],
        [w * 0.5954, h * 0.2815],
      ]
    case "moon":
      return [[w, 0], L, [w, h], [w - w * a("adj", 50000), h / 2]]
    case "cloud":
      return [
        [w * 0.9992, h / 2],
        [w / 2, h * 0.9989],
        [w * 0.0031, h / 2],
        [w / 2, h * 0.0572],
      ]
    case "cloudCallout":
      return [[w * 0.0031, h / 2], [w / 2, h * 0.9989], [w * 0.9992, h / 2], [w / 2, h * 0.0572], tip()]
    case "rightArrow":
    case "leftArrow": {
      const head = Math.min(ss * a("adj2", 50000), w)
      const x = name === "rightArrow" ? w - head : head
      return [[x, 0], L, [x, h], R]
    }
    case "leftRightArrow": {
      const shaft = h * a("adj1", 50000)
      const head = Math.min(ss * a("adj2", 50000), w / 2)
      const y0 = (h - shaft) / 2
      return [[w - head, 0], [w / 2, y0], [head, 0], L, [head, h], [w / 2, y0 + shaft], [w - head, h], R]
    }
    case "chevron": {
      const x = Math.min(ss * a("adj", 50000), w)
      return [[(w - x) / 2, 0], [x, h / 2], [(w - x) / 2, h], R]
    }
    case "homePlate": {
      const x = Math.min(ss * a("adj", 50000), w)
      return [[(w - x) / 2, 0], L, [(w - x) / 2, h], R]
    }
    case "star5":
      return corners(5)
    case "star6": {
      // Office's order: from the upper right, clockwise round to the top.
      const c = corners(6)
      return [c[5], c[4], c[3], c[2], c[1], c[0]]
    }
    default:
      // Rectangles, the diamond, the cross, the frame, the sun, up and down arrows, flowcharts, the 4-point star.
      return [T, L, B, R]
  }
}

const inset = (w: number, h: number, k: number) => ({ x: w * k, y: h * k, w: w * (1 - 2 * k), h: h * (1 - 2 * k) })

/** Insert › Shapes, by group, as Excel's gallery lists them (Cells' subset). */
export const SHAPE_GALLERY: { group: string; shapes: string[] }[] = [
  { group: "Lines", shapes: ["line", "straightConnector1", "freeform", "scribble"] },
  {
    group: "Rectangles",
    shapes: ["rect", "roundRect", "snip1Rect"],
  },
  {
    group: "Basic Shapes",
    shapes: [
      "ellipse",
      "triangle",
      "rtTriangle",
      "parallelogram",
      "trapezoid",
      "diamond",
      "pentagon",
      "hexagon",
      "octagon",
      "plus",
      "can",
      "donut",
      "frame",
      "heart",
      "smileyFace",
      "lightningBolt",
      "sun",
      "moon",
      "cloud",
    ],
  },
  {
    group: "Block Arrows",
    shapes: ["rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "chevron", "homePlate"],
  },
  {
    group: "Flowchart",
    shapes: ["flowChartProcess", "flowChartDecision", "flowChartTerminator", "flowChartAlternateProcess"],
  },
  { group: "Stars and Banners", shapes: ["star4", "star5", "star6"] },
  { group: "Callouts", shapes: ["wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout"] },
]
