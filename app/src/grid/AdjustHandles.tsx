import { useState } from "react"
import type { SheetShape } from "@cells/core"
import { adjustHandles, presetGeometry } from "@/drawing/geometry"
import type { SheetController } from "@/sheet/controller"

/**
 * A selected shape's adjustment handles (Excel's yellow dots): dragging one reshapes the shape (a rounded corner's
 * radius, an arrow's head, a callout's point), with its outline showing the new shape until the drag ends, which is
 * one undo step. The handles sit in the shape's frame, so they turn and flip with it.
 */
export function AdjustHandles({
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
  const [preview, setPreview] = useState<Record<string, number> | null>(null)
  const adjust = { ...shape.adjust, ...preview }
  const handles = adjustHandles(shape.geometry, box.w, box.h, adjust)
  if (!handles.length) return null
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

  const outline = preview ? presetGeometry(shape.geometry, box.w, box.h, adjust) : null
  const flips = shape.flipH || shape.flipV ? `scale(${shape.flipH ? -1 : 1} ${shape.flipV ? -1 : 1})` : ""
  return (
    <>
      {outline ? (
        <svg
          className="adjust-ghost"
          width={box.w * zoom}
          height={box.h * zoom}
          viewBox={`0 0 ${box.w} ${box.h}`}
          aria-hidden
        >
          <g
            transform={flips ? `translate(${shape.flipH ? box.w : 0} ${shape.flipV ? box.h : 0}) ${flips}` : undefined}
          >
            {outline.paths.map((d, i) => (
              <path key={i} d={d} vectorEffect="non-scaling-stroke" />
            ))}
          </g>
        </svg>
      ) : null}
      {handles.map((h, i) => (
        <span
          key={i}
          className="adjust-handle"
          title="Adjust the shape"
          aria-label="Adjustment handle"
          style={{ left: flipX(h.x) * zoom, top: flipY(h.y) * zoom }}
          onPointerDown={(event) => {
            if (event.button !== 0) return
            event.stopPropagation()
            event.preventDefault()
            try {
              ;(event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
            } catch {
              // The pointer is gone already; the drag still works while it's over the handle.
            }
            setPreview({})
          }}
          onPointerMove={(event) => {
            if (!preview) return
            const p = local(event)
            // Each drag reads the handle from the shape as it was, so its other values stay.
            const from = adjustHandles(shape.geometry, box.w, box.h, { ...shape.adjust, ...preview })[i]
            setPreview({ ...preview, ...from.drag(p.x, p.y) })
          }}
          onPointerUp={() => {
            const next = preview
            setPreview(null)
            if (next && Object.keys(next).length)
              objects.update(
                "Adjust Shape",
                (o) => {
                  if (o.id === shape.id) (o as SheetShape).adjust = { ...shape.adjust, ...next }
                },
                [shape.id],
              )
          }}
          onPointerCancel={() => setPreview(null)}
        />
      ))}
    </>
  )
}
