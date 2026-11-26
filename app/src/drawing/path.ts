import type { PathCommand, ShapePath } from "@cells/core"
import { presetGeometry, type Geometry } from "./geometry"

/**
 * Shapes' own outlines (Excel's custGeom): drawn as SVG at a size, made from a preset's SVG paths (Edit Points turns
 * a preset into its outline, as Excel does), and their points for Edit Points.
 */

const f = (n: number) => Math.round(n * 100) / 100

/** The outline at w × h pixels as SVG path data. */
export function pathSvg(p: ShapePath, w: number, h: number): string {
  const sx = p.w ? w / p.w : 1
  const sy = p.h ? h / p.h : 1
  return p.commands
    .map((c) =>
      c.t === "Z"
        ? "Z"
        : c.t === "C"
          ? `C${f(c.x1 * sx)} ${f(c.y1 * sy)} ${f(c.x2 * sx)} ${f(c.y2 * sy)} ${f(c.x * sx)} ${f(c.y * sy)}`
          : `${c.t}${f(c.x * sx)} ${f(c.y * sy)}`,
    )
    .join(" ")
}

/** An arc (SVG's A) as cubic curves, from (x0, y0). */
function arcToCubics(
  x0: number,
  y0: number,
  rx: number,
  ry: number,
  angle: number,
  large: boolean,
  sweep: boolean,
  x: number,
  y: number,
): PathCommand[] {
  if (!rx || !ry || (x0 === x && y0 === y)) return [{ t: "L", x, y }]
  const phi = (angle * Math.PI) / 180
  const cos = Math.cos(phi)
  const sin = Math.sin(phi)
  const dx = (x0 - x) / 2
  const dy = (y0 - y) / 2
  const x1 = cos * dx + sin * dy
  const y1 = -sin * dx + cos * dy
  rx = Math.abs(rx)
  ry = Math.abs(ry)
  const lambda = (x1 * x1) / (rx * rx) + (y1 * y1) / (ry * ry)
  if (lambda > 1) {
    rx *= Math.sqrt(lambda)
    ry *= Math.sqrt(lambda)
  }
  const sign = large === sweep ? -1 : 1
  const num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1
  const coef = sign * Math.sqrt(Math.max(0, num / (rx * rx * y1 * y1 + ry * ry * x1 * x1)))
  const cx1 = (coef * rx * y1) / ry
  const cy1 = (-coef * ry * x1) / rx
  const cx = cos * cx1 - sin * cy1 + (x0 + x) / 2
  const cy = sin * cx1 + cos * cy1 + (y0 + y) / 2
  const ang = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
    return a
  }
  const t1 = ang(1, 0, (x1 - cx1) / rx, (y1 - cy1) / ry)
  let dt = ang((x1 - cx1) / rx, (y1 - cy1) / ry, (-x1 - cx1) / rx, (-y1 - cy1) / ry)
  if (!sweep && dt > 0) dt -= 2 * Math.PI
  if (sweep && dt < 0) dt += 2 * Math.PI
  const parts = Math.ceil(Math.abs(dt) / (Math.PI / 2))
  const step = dt / parts
  const k = (4 / 3) * Math.tan(step / 4)
  const point = (t: number) => ({
    x: cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin,
    y: cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos,
  })
  const deriv = (t: number) => ({
    x: -rx * Math.sin(t) * cos - ry * Math.cos(t) * sin,
    y: -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos,
  })
  const out: PathCommand[] = []
  for (let i = 0; i < parts; i++) {
    const a = t1 + i * step
    const b = a + step
    const p0 = point(a)
    const p3 = point(b)
    const d0 = deriv(a)
    const d3 = deriv(b)
    out.push({
      t: "C",
      x1: p0.x + k * d0.x,
      y1: p0.y + k * d0.y,
      x2: p3.x - k * d3.x,
      y2: p3.y - k * d3.y,
      x: i === parts - 1 ? x : p3.x,
      y: i === parts - 1 ? y : p3.y,
    })
  }
  return out
}

/** SVG path data (absolute commands, as Cells' presets write them) as an outline's commands. */
export function parseSvgPath(d: string): PathCommand[] {
  const tokens = d.match(/[MLHVCQAZ]|-?\d*\.?\d+(?:e-?\d+)?/gi) ?? []
  const out: PathCommand[] = []
  let i = 0
  let cmd = ""
  let x = 0
  let y = 0
  let startX = 0
  let startY = 0
  const num = () => Number(tokens[i++])
  while (i < tokens.length) {
    if (/^[a-z]$/i.test(tokens[i])) cmd = tokens[i++].toUpperCase()
    switch (cmd) {
      case "M":
        x = num()
        y = num()
        startX = x
        startY = y
        out.push({ t: "M", x, y })
        cmd = "L"
        break
      case "L":
        x = num()
        y = num()
        out.push({ t: "L", x, y })
        break
      case "H":
        x = num()
        out.push({ t: "L", x, y })
        break
      case "V":
        y = num()
        out.push({ t: "L", x, y })
        break
      case "C": {
        const c = { t: "C" as const, x1: num(), y1: num(), x2: num(), y2: num(), x: num(), y: num() }
        out.push(c)
        x = c.x
        y = c.y
        break
      }
      case "Q": {
        const qx = num()
        const qy = num()
        const ex = num()
        const ey = num()
        out.push({
          t: "C",
          x1: x + (2 / 3) * (qx - x),
          y1: y + (2 / 3) * (qy - y),
          x2: ex + (2 / 3) * (qx - ex),
          y2: ey + (2 / 3) * (qy - ey),
          x: ex,
          y: ey,
        })
        x = ex
        y = ey
        break
      }
      case "A": {
        const rx = num()
        const ry = num()
        const rot = num()
        const large = num() !== 0
        const sweep = num() !== 0
        const ex = num()
        const ey = num()
        out.push(...arcToCubics(x, y, rx, ry, rot, large, sweep, ex, ey))
        x = ex
        y = ey
        break
      }
      case "Z":
        out.push({ t: "Z" })
        x = startX
        y = startY
        break
      default:
        // Something Cells' presets don't write: skip the token.
        i++
    }
  }
  return out
}

/** A point of an outline that Edit Points shows (each command's end; a close has none). */
export type PathVertex = { path: number; command: number; x: number; y: number }

/** The outline's points at w × h pixels. */
export function pathVertices(paths: ShapePath[], w: number, h: number): PathVertex[] {
  const out: PathVertex[] = []
  paths.forEach((p, pi) => {
    const sx = p.w ? w / p.w : 1
    const sy = p.h ? h / p.h : 1
    p.commands.forEach((c, ci) => {
      if (c.t === "Z") return
      // A figure closed back on its start: the closing point is the start's (shown once).
      const next = p.commands[ci + 1]
      const start = p.commands
        .slice(0, ci)
        .reverse()
        .find((x) => x.t === "M")
      if (next?.t === "Z" && start && start.t === "M" && start.x === c.x && start.y === c.y) return
      out.push({ path: pi, command: ci, x: c.x * sx, y: c.y * sy })
    })
  })
  return out
}

/** Every point of the outlines (ends and curve handles) in pixels at w × h, as one list to measure. */
function allPoints(paths: ShapePath[], w: number, h: number): { x: number; y: number }[] {
  const pts: { x: number; y: number }[] = []
  for (const p of paths) {
    const sx = p.w ? w / p.w : 1
    const sy = p.h ? h / p.h : 1
    for (const c of p.commands) {
      if (c.t === "Z") continue
      if (c.t === "C") pts.push({ x: c.x1 * sx, y: c.y1 * sy }, { x: c.x2 * sx, y: c.y2 * sy })
      pts.push({ x: c.x * sx, y: c.y * sy })
    }
  }
  return pts
}

/**
 * Outlines in pixels at w × h moved so their points fit a new box exactly (Excel's box follows the points): the box
 * (in the old box's pixels) and the outlines in its own pixels.
 */
export function fitPaths(
  paths: ShapePath[],
  w: number,
  h: number,
): { box: { x: number; y: number; w: number; h: number }; paths: ShapePath[] } {
  const pts = allPoints(paths, w, h)
  if (!pts.length) return { box: { x: 0, y: 0, w, h }, paths }
  const x0 = Math.min(...pts.map((p) => p.x))
  const y0 = Math.min(...pts.map((p) => p.y))
  const bw = Math.max(1, Math.max(...pts.map((p) => p.x)) - x0)
  const bh = Math.max(1, Math.max(...pts.map((p) => p.y)) - y0)
  const moved = paths.map((p) => {
    const sx = p.w ? w / p.w : 1
    const sy = p.h ? h / p.h : 1
    const X = (v: number) => v * sx - x0
    const Y = (v: number) => v * sy - y0
    return {
      ...p,
      w: bw,
      h: bh,
      commands: p.commands.map((c): PathCommand =>
        c.t === "Z"
          ? c
          : c.t === "C"
            ? { t: "C", x1: X(c.x1), y1: Y(c.y1), x2: X(c.x2), y2: Y(c.y2), x: X(c.x), y: Y(c.y) }
            : { t: c.t, x: X(c.x), y: Y(c.y) },
      ),
    }
  })
  return { box: { x: x0, y: y0, w: bw, h: bh }, paths: moved }
}

/**
 * An outline's point moved to (x, y) in pixels at w × h; a curve's handles next to it move along, so its curves keep
 * their shape near the point. A figure's start moves its closing point too.
 */
export function movePoint(
  paths: ShapePath[],
  v: { path: number; command: number },
  x: number,
  y: number,
  w: number,
  h: number,
): ShapePath[] {
  return paths.map((p, pi) => {
    if (pi !== v.path) return p
    const sx = p.w ? w / p.w : 1
    const sy = p.h ? h / p.h : 1
    const target = p.commands[v.command]
    if (!target || target.t === "Z") return p
    const nx = x / sx
    const ny = y / sy
    const dx = nx - target.x
    const dy = ny - target.y
    const commands = p.commands.map((c) => ({ ...c }))
    const at = (i: number) => commands[i]
    const shift = (i: number) => {
      const c = at(i)
      if (!c || c.t === "Z") return
      c.x += dx
      c.y += dy
      if (c.t === "C") {
        c.x2 += dx
        c.y2 += dy
      }
      const next = at(i + 1)
      if (next?.t === "C") {
        next.x1 += dx
        next.y1 += dy
      }
    }
    shift(v.command)
    // A figure's start is also where its closing segment ends.
    if (target.t === "M") {
      const close = commands.findIndex((c, i) => i > v.command && c.t === "Z")
      const last = close > 0 ? at(close - 1) : undefined
      if (last && last.t !== "Z" && last.t !== "M" && last.x - dx === target.x && last.y - dy === target.y) {
        last.x += dx
        last.y += dy
        if (last.t === "C") {
          last.x2 += dx
          last.y2 += dy
        }
      }
    }
    return { ...p, commands }
  })
}

/** An outline's point taken out (Ctrl+click in Edit Points): its segment joins its neighbours. A figure keeps two. */
export function deletePoint(paths: ShapePath[], v: { path: number; command: number }): ShapePath[] {
  return paths.map((p, pi) => {
    if (pi !== v.path) return p
    const ends = p.commands.filter((c) => c.t !== "Z").length
    if (ends <= 2) return p
    const commands = p.commands.map((c) => ({ ...c }))
    const c = commands[v.command]
    if (!c || c.t === "Z") return p
    if (c.t === "M") {
      // The next point starts the figure.
      const next = commands[v.command + 1]
      if (!next || next.t === "Z") return p
      commands.splice(v.command, 2, { t: "M", x: next.x, y: next.y })
    } else commands.splice(v.command, 1)
    return { ...p, commands }
  })
}

/** A shape's outline at w × h: its own (custom geometry) or its preset's. */
export function shapeGeometry(
  o: { geometry: string; custom?: ShapePath[]; adjust?: Record<string, number> },
  w: number,
  h: number,
): Geometry {
  if (o.geometry !== "custom" || !o.custom?.length) return presetGeometry(o.geometry, w, h, o.adjust)
  const filled = o.custom.filter((p) => p.fill !== false)
  const open = o.custom.filter((p) => p.fill === false)
  return {
    paths: filled.map((p) => pathSvg(p, w, h)),
    lines: open.length ? open.map((p) => pathSvg(p, w, h)) : undefined,
    text: { x: 0, y: 0, w, h },
    open: !filled.length || undefined,
  }
}

/**
 * Edit Points on a preset: the preset's outline at w × h as the shape's own (Excel turns the shape into a freeform
 * when its points are first edited). Lines a preset draws unfilled (a can's rim) stay unfilled.
 */
export function presetAsPaths(geometry: string, w: number, h: number, adjust?: Record<string, number>): ShapePath[] {
  const g = presetGeometry(geometry, w, h, adjust)
  const out: ShapePath[] = g.paths.map((d) => ({ w, h, commands: parseSvgPath(d) }))
  for (const d of g.lines ?? []) out.push({ w, h, commands: parseSvgPath(d), fill: false })
  if (g.open) for (const p of out) p.fill = false
  return out
}
