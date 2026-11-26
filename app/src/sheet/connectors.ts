import { isChart, type SheetObject, type SheetShape } from "@cells/core"
import type { ChartBox } from "./charts"
import { connectionSites } from "@/drawing/geometry"

/**
 * Connectors glued to shapes (Excel's lines and arrows drawn from one shape to another): where a shape's connection
 * sites are on the sheet, where a line's ends are, and lines re-routed when the shapes they're glued to move, turn or
 * change size. Positions are sheet pixels at 100%.
 */

export type Point = { x: number; y: number }

/** The presets drawn from end to end (Insert › Shapes › Lines). */
export function isLineGeometry(geometry: string): boolean {
  return geometry === "line" || geometry.includes("Connector")
}

/** A line or arrow that can be glued (top-level connectors). */
export function isConnector(o: SheetObject): o is SheetShape {
  return !isChart(o) && o.kind === "shape" && !!o.connector
}

/** A shape's sites on the sheet: its own sites flipped and turned as the shape is. */
export function sitesOf(o: SheetShape, box: ChartBox): Point[] {
  if (o.connector || o.ink) return []
  const turn = ((o.rotation ?? 0) * Math.PI) / 180
  const cx = box.w / 2
  const cy = box.h / 2
  return connectionSites(o.geometry, box.w, box.h, o.adjust).map(([x, y]) => {
    const fx = (o.flipH ? box.w - x : x) - cx
    const fy = (o.flipV ? box.h - y : y) - cy
    return {
      x: box.x + cx + fx * Math.cos(turn) - fy * Math.sin(turn),
      y: box.y + cy + fx * Math.sin(turn) + fy * Math.cos(turn),
    }
  })
}

/** A line's ends: from its box and flips (a line goes from the top left to the bottom right unless flipped). */
export function lineEnds(o: SheetShape, box: ChartBox): [Point, Point] {
  const x0 = o.flipH ? box.x + box.w : box.x
  const x1 = o.flipH ? box.x : box.x + box.w
  const y0 = o.flipV ? box.y + box.h : box.y
  const y1 = o.flipV ? box.y : box.y + box.h
  return [
    { x: x0, y: y0 },
    { x: x1, y: y1 },
  ]
}

/** The box and flips of a line from `a` to `b`. */
export function lineBox(a: Point, b: Point): { box: ChartBox; flipH: boolean; flipV: boolean } {
  return {
    box: { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y) },
    flipH: b.x < a.x,
    flipV: b.y < a.y,
  }
}

/** The site nearest a point within `radius` pixels, on any shape but `skip` (a line's own). */
export function nearestSite(
  objects: readonly SheetObject[],
  at: Point,
  radius: number,
  boxOf: (o: SheetObject) => ChartBox,
  skip?: number,
): { id: number; site: number; point: Point } | null {
  let best: { id: number; site: number; point: Point; d: number } | null = null
  for (const o of objects) {
    if (isChart(o) || o.kind !== "shape" || o.connector || o.hidden || o.id === skip) continue
    sitesOf(o, boxOf(o)).forEach((p, site) => {
      const d = Math.hypot(p.x - at.x, p.y - at.y)
      if (d <= radius && (!best || d < best.d)) best = { id: o.id, site, point: p, d }
    })
  }
  if (!best) return null
  const { id, site, point } = best
  return { id, site, point }
}

/** The shape under a point (the topmost), whose sites show while a line is drawn or its end dragged. */
export function shapeAt(
  objects: readonly SheetObject[],
  at: Point,
  boxOf: (o: SheetObject) => ChartBox,
  margin = 0,
): SheetShape | null {
  for (let i = objects.length - 1; i >= 0; i--) {
    const o = objects[i]
    if (isChart(o) || o.kind !== "shape" || o.connector || o.hidden) continue
    const b = boxOf(o)
    if (at.x >= b.x - margin && at.x <= b.x + b.w + margin && at.y >= b.y - margin && at.y <= b.y + b.h + margin)
      return o
  }
  return null
}

/**
 * The drawing with each glued connector's ends on its shapes' sites again (connectors unchanged stay the same
 * objects). A glued shape that's gone, or a site it no longer has, lets go of that end.
 */
export function routeConnectors(
  objects: SheetObject[],
  boxOf: (o: SheetObject) => ChartBox,
  place: (o: SheetShape, box: ChartBox) => SheetShape,
): SheetObject[] {
  if (!objects.some((o) => isConnector(o) && (o.from || o.to))) return objects
  const byId = new Map(objects.map((o) => [o.id, o]))
  const siteAt = (glue: SheetShape["from"]): Point | null => {
    const target = glue && byId.get(glue.id)
    if (!target || isChart(target) || target.kind !== "shape" || target.connector) return null
    return sitesOf(target, boxOf(target))[glue.site] ?? null
  }
  return objects.map((o) => {
    if (!isConnector(o) || (!o.from && !o.to)) return o
    const box = boxOf(o)
    const [a, b] = lineEnds(o, box)
    const from = o.from ? siteAt(o.from) : null
    const to = o.to ? siteAt(o.to) : null
    const start = from ?? a
    const end = to ?? b
    const lost = (!!o.from && !from) || (!!o.to && !to)
    const near = (p: Point, q: Point) => Math.abs(p.x - q.x) < 0.5 && Math.abs(p.y - q.y) < 0.5
    if (!lost && near(start, a) && near(end, b)) return o
    const next = lineBox(start, end)
    // A connector runs straight from end to end (a turn it had goes).
    const moved: SheetShape = { ...o, flipH: next.flipH, flipV: next.flipV }
    if (!moved.flipH) delete moved.flipH
    if (!moved.flipV) delete moved.flipV
    delete moved.rotation
    if (o.from && !from) delete moved.from
    if (o.to && !to) delete moved.to
    return place(moved, next.box)
  })
}
