import { useRef } from "react"

/**
 * A thin scrollbar drawn by the app: the sheet is far taller than a browser can scroll natively, and its range
 * follows the used area (plus where the view is), as Excel's does.
 */
export function Scrollbar({
  orientation,
  length,
  offset,
  position,
  view,
  content,
  onScroll,
  onDone,
}: {
  orientation: "vertical" | "horizontal"
  /** Length of the whole bar (the canvas side it runs along). */
  length: number
  /** The header size it starts after. */
  offset: number
  position: number
  view: number
  content: number
  onScroll: (position: number) => void
  onDone: () => void
}) {
  const trackRef = useRef<HTMLDivElement>(null)
  const dragRef = useRef<{ start: number; position: number } | null>(null)
  const track = Math.max(0, length - offset)
  const thumb = Math.max(18, Math.min(track, (view / Math.max(view, content)) * track))
  const travel = Math.max(1, track - thumb)
  const maxPosition = Math.max(1, content - view)
  const at = Math.min(travel, (position / maxPosition) * travel)
  const vertical = orientation === "vertical"

  const coordinate = (event: React.PointerEvent) => (vertical ? event.clientY : event.clientX)

  return (
    <div
      ref={trackRef}
      className={`grid-scrollbar is-${orientation}`}
      style={vertical ? { top: offset, height: track } : { left: offset, width: track }}
      onPointerDown={(event) => {
        if (event.target !== trackRef.current) return
        // A click on the track pages towards the click.
        const rect = trackRef.current!.getBoundingClientRect()
        const point = coordinate(event) - (vertical ? rect.top : rect.left)
        onScroll(Math.max(0, position + (point < at ? -view : view)))
        onDone()
      }}
      aria-hidden
    >
      <div
        className="grid-scrollbar-thumb"
        style={vertical ? { top: at, height: thumb } : { left: at, width: thumb }}
        onPointerDown={(event) => {
          event.stopPropagation()
          event.currentTarget.setPointerCapture(event.pointerId)
          dragRef.current = { start: coordinate(event), position }
        }}
        onPointerMove={(event) => {
          const drag = dragRef.current
          if (!drag) return
          const delta = coordinate(event) - drag.start
          onScroll(Math.max(0, drag.position + (delta / travel) * maxPosition))
        }}
        onPointerUp={() => {
          dragRef.current = null
          onDone()
        }}
      />
    </div>
  )
}
