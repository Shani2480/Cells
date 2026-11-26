import { useEffect, useRef, useState } from "react"
import type { Point } from "@/sheet/connectors"

/** The gallery's two freehand tools (Insert › Shapes › Lines). */
export const isFreeformTool = (geometry: string) => geometry === "freeform" || geometry === "scribble"

/** How near (screen pixels) the pointer must come back to the start to close the shape. */
const CLOSE_RADIUS = 8

/**
 * Drawing a freeform, as Excel does. Freeform: Shape: each click adds a corner and a drag draws freehand; clicking
 * the start again closes it into a shape, a double-click or Enter ends it as a line (Escape too). Freeform: Scribble:
 * one stroke drawn while the button is down, closed into a shape when it ends at its start.
 */
export function FreeformDraw({
  geometry,
  zoom,
  sheetPoint,
  toLayer,
  onFinish,
  onCancel,
}: {
  geometry: string
  zoom: number
  sheetPoint: (event: { clientX: number; clientY: number }) => Point
  toLayer: (p: Point) => Point
  onFinish: (points: Point[], closed: boolean) => void
  onCancel: () => void
}) {
  const scribble = geometry === "scribble"
  const [points, setPoints] = useState<Point[]>([])
  const [cursor, setCursor] = useState<Point | null>(null)
  const down = useRef(false)
  const now = useRef(points)
  now.current = points

  const nearStart = (p: Point, list = now.current) =>
    list.length > 2 && Math.hypot(p.x - list[0].x, p.y - list[0].y) * zoom <= CLOSE_RADIUS
  const finish = (closed: boolean) => {
    const list = now.current
    down.current = false
    setPoints([])
    if (list.length < 2) onCancel()
    else onFinish(closed ? [...list] : list, closed)
  }

  // Enter or Escape ends a freeform as a line (Escape with nothing drawn stops drawing).
  useEffect(() => {
    if (scribble) return
    const key = (e: KeyboardEvent) => {
      if (e.key !== "Enter" && e.key !== "Escape") return
      e.preventDefault()
      e.stopPropagation()
      finish(false)
    }
    window.addEventListener("keydown", key, true)
    return () => window.removeEventListener("keydown", key, true)
  }, [scribble])

  const shown = cursor && !scribble && points.length && !down.current ? [...points, cursor] : points
  const d = shown.map((p, i) => `${i ? "L" : "M"}${toLayer(p).x} ${toLayer(p).y}`).join("")
  const start = points[0] ? toLayer(points[0]) : null
  return (
    <div
      className="object-draw freeform-draw"
      onPointerDown={(event) => {
        if (event.button !== 0) return
        event.preventDefault()
        event.stopPropagation()
        try {
          ;(event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
        } catch {
          // The pointer is gone: the stroke still draws while it's over the sheet.
        }
        const at = sheetPoint(event)
        if (!scribble && nearStart(at)) return finish(true)
        down.current = true
        setPoints(scribble ? [at] : [...now.current, at])
      }}
      onPointerMove={(event) => {
        const at = sheetPoint(event)
        setCursor(at)
        if (down.current) setPoints([...now.current, at])
      }}
      onPointerUp={(event) => {
        if (!down.current) return
        down.current = false
        if (scribble) finish(nearStart(sheetPoint(event)))
      }}
      onDoubleClick={(event) => {
        event.preventDefault()
        event.stopPropagation()
        if (!scribble) finish(false)
      }}
    >
      {d ? (
        <svg className="freeform-ghost" aria-hidden>
          <path d={d} />
          {start && !scribble ? (
            <circle
              cx={start.x}
              cy={start.y}
              r={4}
              className={cursor && nearStart(cursor) ? "is-closing" : undefined}
            />
          ) : null}
        </svg>
      ) : null}
    </div>
  )
}
