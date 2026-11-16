import { useEffect, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"

/** Dragged down this far, or flicked down this fast (px per ms), the sheet closes. */
export const DISMISS_DISTANCE = 96
export const DISMISS_SPEED = 0.5
/** A touch on the header that moves this far (CSS px) is a drag, not a tap. */
export const DRAG_SLOP = 8

/** Whether a drag on the handle, ending this far down at this speed, closes the sheet. */
export function dismisses(distance: number, speed: number): boolean {
  return distance > DISMISS_DISTANCE || (distance > 16 && speed > DISMISS_SPEED)
}

const reducedMotion = () => typeof matchMedia === "function" && matchMedia("(prefers-reduced-motion: reduce)").matches
const SLIDE_MS = 220

/**
 * The phone's panel from the bottom edge (Excel for phones' command palette, the sheet list): it slides up over a
 * dimmed sheet, follows a finger dragging its handle down and closes past a threshold or on a flick, or on a tap
 * outside, Back (Esc) or `onClose`. The sheet is still there under it: tapping outside closes, it doesn't select.
 */
export function BottomSheet({
  label,
  header,
  onClose,
  children,
}: {
  label: string
  header: ReactNode
  onClose: () => void
  children: ReactNode
}) {
  const panelRef = useRef<HTMLDivElement>(null)
  const [shown, setShown] = useState(false)
  const [offset, setOffset] = useState(0)
  const [closing, setClosing] = useState(false)
  const drag = useRef<{
    y: number
    t: number
    lastY: number
    lastT: number
    speed: number
    dragging: boolean
  } | null>(null)
  const dragged = useRef(false)

  // Focus moves into the sheet, so Back's Esc (android-files.ts) reaches it rather than the grid, which takes Esc for
  // itself, and the on-screen keyboard goes down.
  useEffect(() => {
    panelRef.current?.focus({ preventScroll: true })
  }, [])

  // Slide in on the frame after the first, so the start position is drawn first.
  useEffect(() => {
    const frame = requestAnimationFrame(() => setShown(true))
    return () => cancelAnimationFrame(frame)
  }, [])

  const close = () => {
    if (closing) return
    setClosing(true)
    if (reducedMotion()) onClose()
    else window.setTimeout(onClose, SLIDE_MS)
  }
  const closeRef = useRef(close)
  closeRef.current = close

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) return
      // A dialog over the sheet takes its own Esc.
      if (document.querySelector('.dialog[aria-modal="true"]')) return
      event.preventDefault()
      closeRef.current()
    }
    window.addEventListener("keydown", onKey)
    return () => window.removeEventListener("keydown", onKey)
  }, [])

  // A touch anywhere on the header, its buttons too: it becomes a drag once it has moved down a little, and then the
  // tap that ends it doesn't press the button it started on.
  const onPointerDown = (event: React.PointerEvent) => {
    const now = performance.now()
    drag.current = { y: event.clientY, t: now, lastY: event.clientY, lastT: now, speed: 0, dragging: false }
  }
  const onPointerMove = (event: React.PointerEvent) => {
    const d = drag.current
    if (!d) return
    const dy = event.clientY - d.y
    if (!d.dragging) {
      if (Math.abs(dy) < DRAG_SLOP) return
      d.dragging = true
      event.currentTarget.setPointerCapture?.(event.pointerId)
    }
    const now = performance.now()
    if (now > d.lastT) d.speed = (event.clientY - d.lastY) / (now - d.lastT)
    d.lastY = event.clientY
    d.lastT = now
    // Up past the open position only stretches a little.
    setOffset(dy > 0 ? dy : dy / 4)
  }
  const onPointerUp = (event: React.PointerEvent) => {
    const d = drag.current
    drag.current = null
    if (!d?.dragging) return
    dragged.current = true
    const dy = event.clientY - d.y
    if (dismisses(dy, d.speed)) close()
    setOffset(0)
  }
  // The click that follows a drag.
  const onClickCapture = (event: React.MouseEvent) => {
    if (!dragged.current) return
    dragged.current = false
    event.preventDefault()
    event.stopPropagation()
  }

  const open = shown && !closing
  return createPortal(
    <div className="phone-sheet-layer" data-open={open || undefined}>
      <div className="phone-sheet-backdrop" onPointerDown={(event) => (event.preventDefault(), close())} />
      <div
        ref={panelRef}
        tabIndex={-1}
        className="phone-sheet"
        role="dialog"
        aria-label={label}
        data-dragging={drag.current?.dragging ? "" : undefined}
        style={offset ? { translate: `0 ${Math.max(offset, -24)}px` } : undefined}
      >
        <div
          className="phone-sheet-grip"
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onPointerCancel={onPointerUp}
          onClickCapture={onClickCapture}
        >
          <span className="phone-sheet-handle" aria-hidden />
          {header}
        </div>
        <div className="phone-sheet-body">{children}</div>
      </div>
    </div>,
    document.body,
  )
}
