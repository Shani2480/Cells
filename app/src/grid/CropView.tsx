import type { SheetPicture, Theme } from "@cells/core"
import { useRef, useState } from "react"
import { cropFrame, cropOf, flipCrop, fullBox, scaleFull, type Box, type CropHandle } from "@/drawing/crop"
import { canvasMeasure, objectDocument } from "@/drawing/render"
import type { ObjectCommands } from "@/sheet/objects"

const CROP_HANDLES: CropHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"]
const CORNERS = ["nw", "ne", "se", "sw"] as const
/** A drag starts after this many pixels, so a click only selects. */
const DRAG_START = 3

type Drag = {
  how: CropHandle | "pan" | `full:${(typeof CORNERS)[number]}`
  x: number
  y: number
  frame: Box
  full: Box
  moved: boolean
}

/**
 * Picture Format › Crop, on the sheet: the whole picture shows faintly around the frame, which keeps the part shown.
 * The black handles crop (the frame's sides move, the picture stays; Shift on a corner keeps the frame's
 * proportions), the round ones size the whole picture, and a drag anywhere else slides the picture under the frame.
 * Each drag is one undo step; Esc, Enter, a double-click or a click away ends cropping.
 */
export function CropView({
  picture: o,
  frame: shownFrame,
  objects,
  theme,
  zoom,
  screen,
  onDone,
}: {
  picture: SheetPicture
  frame: Box
  objects: ObjectCommands
  theme: Theme
  zoom: number
  /** Where a box (sheet pixels) is on the layer. */
  screen: (box: Box) => { left: number; top: number; width: number; height: number }
  onDone: () => void
}) {
  const [live, setLiveState] = useState<{ frame: Box; full: Box } | null>(null)
  const liveNow = useRef(live)
  const setLive = (v: typeof live) => {
    liveNow.current = v
    setLiveState(v)
  }
  const drag = useRef<Drag | null>(null)

  const seen = objects.seenCrop(o)
  const frame = live?.frame ?? shownFrame
  const full = live?.full ?? fullBox(shownFrame, seen)
  const rotation = o.rotation ?? 0
  const flips = o.flipH || o.flipV ? `scale(${o.flipH ? -1 : 1}, ${o.flipV ? -1 : 1})` : undefined
  const shown = { ...o, crop: flipCrop(cropOf(frame, full), o.flipH, o.flipV) }
  const svg = objectDocument(shown, frame.w, frame.h, theme, canvasMeasure)

  const begin = (event: React.PointerEvent, how: Drag["how"]) => {
    if (event.button !== 0) return
    event.stopPropagation()
    event.preventDefault()
    drag.current = { how, x: event.clientX, y: event.clientY, frame, full, moved: false }
    try {
      ;(event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
    } catch {
      // A pointer that's already gone: the drag still works while it's over the picture.
    }
  }
  const move = (event: React.PointerEvent) => {
    const d = drag.current
    if (!d) return
    const sx = (event.clientX - d.x) / zoom
    const sy = (event.clientY - d.y) / zoom
    if (!d.moved && Math.hypot(sx, sy) < DRAG_START / zoom) return
    d.moved = true
    // A turned picture crops along its own sides.
    const a = (-rotation * Math.PI) / 180
    const dx = sx * Math.cos(a) - sy * Math.sin(a)
    const dy = sx * Math.sin(a) + sy * Math.cos(a)
    if (d.how === "pan") setLive({ frame: d.frame, full: { ...d.full, x: d.full.x + dx, y: d.full.y + dy } })
    else if (d.how.startsWith("full:"))
      setLive({ frame: d.frame, full: scaleFull(d.full, d.how.slice(5) as (typeof CORNERS)[number], dx, dy) })
    else setLive({ frame: cropFrame(d.frame, d.how as CropHandle, dx, dy, event.shiftKey), full: d.full })
  }
  const end = () => {
    const d = drag.current
    const last = liveNow.current
    drag.current = null
    if (d?.moved && last) {
      const label = d.how === "pan" ? "Move Picture" : d.how.startsWith("full:") ? "Size Picture" : "Crop"
      objects.setCrop(o.id, cropOf(last.frame, last.full), last.frame, label)
    }
    setLive(null)
    onDone()
  }
  const events = {
    onPointerMove: move,
    onPointerUp: end,
    onPointerCancel: () => ((drag.current = null), setLive(null)),
  }

  const at = (b: Box) => ({
    left: (b.x - frame.x) * zoom,
    top: (b.y - frame.y) * zoom,
    width: b.w * zoom,
    height: b.h * zoom,
  })
  return (
    <div
      className="crop-view"
      style={{ ...screen(frame), transform: rotation ? `rotate(${rotation}deg)` : undefined }}
      data-object={o.id}
      role="img"
      aria-label={`${o.name}, cropping`}
      onDoubleClick={(event) => {
        event.stopPropagation()
        objects.endCrop()
        onDone()
      }}
    >
      <div className="crop-full" style={at(full)} onPointerDown={(e) => begin(e, "pan")} {...events}>
        <img src={o.svg ?? o.src} alt="" draggable={false} style={flips ? { transform: flips } : undefined} />
        {CORNERS.map((c) => (
          <span
            key={c}
            className={`crop-full-handle chart-handle is-${c}`}
            onPointerDown={(e) => begin(e, `full:${c}`)}
            {...events}
          />
        ))}
      </div>
      <div className="crop-frame" onPointerDown={(e) => begin(e, "pan")} {...events}>
        <svg
          className="object-svg"
          width="100%"
          height="100%"
          viewBox={`0 0 ${frame.w} ${frame.h}`}
          preserveAspectRatio="none"
          aria-hidden
          style={flips ? { transform: flips } : undefined}
          dangerouslySetInnerHTML={{ __html: svg }}
        />
        {CROP_HANDLES.map((h) => (
          <span key={h} className={`crop-handle is-${h}`} onPointerDown={(e) => begin(e, h)} {...events} />
        ))}
      </div>
    </div>
  )
}
