import type { SheetController } from "@/sheet/controller"
import { shapeAt, sitesOf, type Point } from "@/sheet/connectors"

/** A line being drawn (from `fixed`), or a line's end being dragged (`end`: which one of which line). */
export type LineDrag = { fixed: Point; to: Point; end?: { id: number; which: "from" | "to" } }

/** Shift while drawing a line: the nearest 45° step. */
export function stepAngle(from: Point, to: Point): Point {
  const dx = to.x - from.x
  const dy = to.y - from.y
  const step = Math.PI / 4
  const angle = Math.round(Math.atan2(dy, dx) / step) * step
  const len = Math.hypot(dx, dy)
  return { x: from.x + Math.cos(angle) * len, y: from.y + Math.sin(angle) * len }
}

/**
 * What's shown while a line is drawn or its end dragged: the line so far, and the connection sites of the shape
 * under the pointer as grey dots, the one the end would glue to marked (Excel's connector guides).
 */
export function LineGuide({
  objects,
  line,
  hover,
  toLayer,
}: {
  objects: SheetController["objects"]
  line: LineDrag | null
  hover: Point | null
  toLayer: (p: Point) => Point
}) {
  const at = line?.to ?? hover
  if (!at) return null
  const boxOf = (o: Parameters<typeof objects.box>[0]) => objects.box(o)
  const skip = line?.end?.id
  const zoomed = toLayer({ x: 1, y: 0 }).x - toLayer({ x: 0, y: 0 }).x
  const target = shapeAt(
    objects.list.filter((o) => o.id !== skip),
    at,
    boxOf,
    16 / Math.max(0.1, zoomed),
  )
  const glue = objects.glueAt(at, skip)
  const end = glue?.point ?? at
  const sites = target ? sitesOf(target, boxOf(target)) : []
  const a = line ? toLayer(line.fixed) : null
  const b = toLayer(end)
  return (
    <svg className="line-guide" aria-hidden>
      {a ? <line x1={a.x} y1={a.y} x2={b.x} y2={b.y} /> : null}
      {sites.map((p, i) => {
        const s = toLayer(p)
        const on = glue && glue.id === target?.id && glue.site === i
        return <circle key={i} cx={s.x} cy={s.y} r={on ? 5 : 3.5} className={on ? "is-on" : undefined} />
      })}
    </svg>
  )
}
