import { useState } from "react"
import type { ShapePath, SheetShape } from "@cells/core"
import { deletePoint, movePoint, pathSvg, pathVertices, type PathVertex } from "@/drawing/path"
import type { SheetController } from "@/sheet/controller"

/**
 * Edit Points on a selected shape: its outline with a black square at each point. Dragging a point reshapes the
 * outline (shown as it moves; one undo step when let go) and Ctrl+click takes a point out. A preset becomes a
 * freeform once a point changes, and the shape's box follows its points. The points sit in the shape's frame, so
 * they turn and flip with it.
 */
export function EditPoints({
  shape,
  box,
  zoom,
  rotation,
  objects,
}: {
  shape: SheetShape
  box: { w: number; h: number }
  zoom: number
  rotation: number
  objects: SheetController["objects"]
}) {
  const [drag, setDrag] = useState<{ vertex: PathVertex; paths: ShapePath[]; start: ShapePath[] } | null>(null)
  const paths = drag?.paths ?? objects.pointsOf(shape)
  const vertices = pathVertices(paths, box.w, box.h)
  const flipX = (x: number) => (shape.flipH ? box.w - x : x)
  const flipY = (y: number) => (shape.flipV ? box.h - y : y)

  /** A pointer as a point in the shape (its own pixels, turned and flipped back). */
  const local = (event: React.PointerEvent) => {
    const frame = (event.currentTarget as HTMLElement).closest(".object-frame")!.getBoundingClientRect()
    const dx = event.clientX - (frame.left + frame.width / 2)
    const dy = event.clientY - (frame.top + frame.height / 2)
    const r = (-rotation * Math.PI) / 180
    const x = (dx * Math.cos(r) - dy * Math.sin(r)) / zoom + box.w / 2
    const y = (dx * Math.sin(r) + dy * Math.cos(r)) / zoom + box.h / 2
    return { x: flipX(x), y: flipY(y) }
  }

  const flips = shape.flipH || shape.flipV ? `scale(${shape.flipH ? -1 : 1} ${shape.flipV ? -1 : 1})` : ""
  return (
    <>
      <svg
        className="edit-points-outline"
        width={box.w * zoom}
        height={box.h * zoom}
        viewBox={`0 0 ${box.w} ${box.h}`}
        aria-hidden
      >
        <g transform={flips ? `translate(${shape.flipH ? box.w : 0} ${shape.flipV ? box.h : 0}) ${flips}` : undefined}>
          {paths.map((p, i) => (
            <path key={i} d={pathSvg(p, box.w, box.h)} vectorEffect="non-scaling-stroke" />
          ))}
        </g>
      </svg>
      {vertices.map((v) => (
        <span
          key={`${v.path}:${v.command}`}
          className="edit-point"
          title="Drag to move the point; Ctrl+click to delete it"
          aria-label="Point"
          style={{ left: flipX(v.x) * zoom, top: flipY(v.y) * zoom }}
          onPointerDown={(event) => {
            if (event.button !== 0) return
            event.stopPropagation()
            event.preventDefault()
            if (event.ctrlKey || event.metaKey) {
              const next = deletePoint(paths, v)
              if (next !== paths && next.some((p, i) => p !== paths[i]))
                objects.setPoints(shape.id, next, "Delete Point")
              return
            }
            try {
              ;(event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
            } catch {
              // The pointer is gone already; the drag still works while it's over the point.
            }
            setDrag({ vertex: v, paths, start: paths })
          }}
          onPointerMove={(event) => {
            if (!drag) return
            const p = local(event)
            setDrag({ ...drag, paths: movePoint(objects.pointsOf(shape), drag.vertex, p.x, p.y, box.w, box.h) })
          }}
          onPointerUp={() => {
            const next = drag
            setDrag(null)
            if (next && next.paths !== next.start) objects.setPoints(shape.id, next.paths)
          }}
          onPointerCancel={() => setDrag(null)}
        />
      ))}
    </>
  )
}
