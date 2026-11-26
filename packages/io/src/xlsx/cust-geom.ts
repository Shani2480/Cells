import type { PathCommand, ShapePath } from "@cells/core"
import { at, kid, kids, type XNode } from "./dom"

/**
 * A shape's own outline (DrawingML's custGeom: freeforms, scribbles, edited points) read as Cells' paths and written
 * back. Paths whose points are formulas (guides) rather than numbers aren't modelled: the shape keeps the file's XML.
 */

const ANGLE = 60000

/** An arc (DrawingML's arcTo: radii, start and swing in 60000ths of a degree) from (x0, y0) as cubic curves. */
function arcCubics(x0: number, y0: number, wR: number, hR: number, st: number, sw: number): PathCommand[] {
  const a0 = (st / ANGLE) * (Math.PI / 180)
  const da = (sw / ANGLE) * (Math.PI / 180)
  const cx = x0 - wR * Math.cos(a0)
  const cy = y0 - hR * Math.sin(a0)
  const parts = Math.max(1, Math.ceil(Math.abs(da) / (Math.PI / 2)))
  const step = da / parts
  const k = (4 / 3) * Math.tan(step / 4)
  const out: PathCommand[] = []
  for (let i = 0; i < parts; i++) {
    const a = a0 + i * step
    const b = a + step
    out.push({
      t: "C",
      x1: cx + wR * (Math.cos(a) - k * Math.sin(a)),
      y1: cy + hR * (Math.sin(a) + k * Math.cos(a)),
      x2: cx + wR * (Math.cos(b) + k * Math.sin(b)),
      y2: cy + hR * (Math.sin(b) - k * Math.cos(b)),
      x: cx + wR * Math.cos(b),
      y: cy + hR * Math.sin(b),
    })
  }
  return out
}

/** The custGeom's paths, or null when any point isn't a plain number. */
export function readCustGeom(geom: XNode | undefined): ShapePath[] | null {
  const list = kids(kid(geom, "pathLst"), "path")
  if (!list.length) return null
  const out: ShapePath[] = []
  for (const p of list) {
    const commands: PathCommand[] = []
    let x = 0
    let y = 0
    let sx = 0
    let sy = 0
    const pts = (n: XNode) => kids(n, "pt").map((pt) => ({ x: Number(at(pt, "x")), y: Number(at(pt, "y")) }))
    for (const c of p.children) {
      const ps = pts(c)
      if (ps.some((q) => !Number.isFinite(q.x) || !Number.isFinite(q.y))) return null
      switch (c.local) {
        case "moveTo":
          ;[{ x, y }] = ps
          sx = x
          sy = y
          commands.push({ t: "M", x, y })
          break
        case "lnTo":
          ;[{ x, y }] = ps
          commands.push({ t: "L", x, y })
          break
        case "cubicBezTo":
          if (ps.length < 3) return null
          commands.push({ t: "C", x1: ps[0].x, y1: ps[0].y, x2: ps[1].x, y2: ps[1].y, x: ps[2].x, y: ps[2].y })
          ;({ x, y } = ps[2])
          break
        case "quadBezTo":
          if (ps.length < 2) return null
          commands.push({
            t: "C",
            x1: x + (2 / 3) * (ps[0].x - x),
            y1: y + (2 / 3) * (ps[0].y - y),
            x2: ps[1].x + (2 / 3) * (ps[0].x - ps[1].x),
            y2: ps[1].y + (2 / 3) * (ps[0].y - ps[1].y),
            x: ps[1].x,
            y: ps[1].y,
          })
          ;({ x, y } = ps[1])
          break
        case "arcTo": {
          const v = ["wR", "hR", "stAng", "swAng"].map((k) => Number(at(c, k)))
          if (v.some((n) => !Number.isFinite(n))) return null
          const arc = arcCubics(x, y, v[0], v[1], v[2], v[3])
          commands.push(...arc)
          const last = arc[arc.length - 1]
          if (last.t === "C") ({ x, y } = last)
          break
        }
        case "close":
          commands.push({ t: "Z" })
          x = sx
          y = sy
          break
        default:
          return null
      }
    }
    const w = Number(at(p, "w") ?? 0)
    const h = Number(at(p, "h") ?? 0)
    const path: ShapePath = { w: Number.isFinite(w) ? w : 0, h: Number.isFinite(h) ? h : 0, commands }
    if (at(p, "fill") === "none") path.fill = false
    out.push(path)
  }
  return out
}

/** Paths as custGeom, each in the shape's size in EMUs (`cx` × `cy`), as Excel writes them. */
export function custGeomXml(paths: ShapePath[], cx: number, cy: number): string {
  const body = paths
    .map((p) => {
      const kx = p.w ? cx / p.w : 1
      const ky = p.h ? cy / p.h : 1
      const pt = (x: number, y: number) => `<a:pt x="${Math.round(x * kx)}" y="${Math.round(y * ky)}"/>`
      const cmds = p.commands
        .map((c) =>
          c.t === "Z"
            ? "<a:close/>"
            : c.t === "C"
              ? `<a:cubicBezTo>${pt(c.x1, c.y1)}${pt(c.x2, c.y2)}${pt(c.x, c.y)}</a:cubicBezTo>`
              : c.t === "M"
                ? `<a:moveTo>${pt(c.x, c.y)}</a:moveTo>`
                : `<a:lnTo>${pt(c.x, c.y)}</a:lnTo>`,
        )
        .join("")
      return `<a:path w="${cx}" h="${cy}"${p.fill === false ? ` fill="none"` : ""}>${cmds}</a:path>`
    })
    .join("")
  return (
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>` +
    `<a:pathLst>${body}</a:pathLst></a:custGeom>`
  )
}
