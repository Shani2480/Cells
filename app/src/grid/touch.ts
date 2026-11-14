/**
 * Touch on the grid (Excel for tablets): one finger pans and a tap selects; a long press opens the context menu;
 * a double tap edits; two fingers pinch to zoom; the round grippers on the selection's corners stretch it.
 * The arithmetic lives here so it can be tested without a touch screen.
 */

/** Hold this long without moving for the context menu (iPadOS and Android both use about half a second). */
export const LONG_PRESS_MS = 500
/** Two taps this close in time (and on the same cell) are a double tap. */
export const DOUBLE_TAP_MS = 350
/** A finger moving this far (CSS px) is a pan, not a tap. */
export const TAP_SLOP = 8
/** The grippers: drawn this big, and hit within a finger-sized circle around them. */
export const GRIPPER_RADIUS = 6
export const GRIPPER_HIT = 22

export type Point = { x: number; y: number }

export const distance = (a: Point, b: Point) => Math.hypot(a.x - b.x, a.y - b.y)
export const midpoint = (a: Point, b: Point): Point => ({ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 })

/**
 * Pinch: the new zoom follows the change in finger distance, and the scroll moves so that the sheet point under
 * the fingers' midpoint stays under it. `mid` is in CSS pixels of the canvas; `origin` is the top-left of the
 * scrolling cell area in sheet pixels (the row header width and the header height).
 */
export function pinch(options: {
  startZoom: number
  startDistance: number
  distance: number
  mid: Point
  scroll: Point
  origin: Point
  zoom: number
  clamp: (zoom: number) => number
}): { zoom: number; scroll: Point } {
  const { startZoom, startDistance, mid, scroll, origin, zoom, clamp } = options
  const next = clamp(startDistance > 0 ? (startZoom * options.distance) / startDistance : startZoom)
  // The sheet point under the midpoint at the current zoom…
  const sx = scroll.x + mid.x / zoom - origin.x
  const sy = scroll.y + mid.y / zoom - origin.y
  // …is under the midpoint again at the new zoom.
  return {
    zoom: next,
    scroll: { x: Math.max(0, sx - mid.x / next + origin.x), y: Math.max(0, sy - mid.y / next + origin.y) },
  }
}

/** Remembers the last tap so the next one can be recognised as a double tap on the same cell. */
export class TapTracker {
  private last: { time: number; row: number; col: number } | null = null

  /** Records a tap; true when it completes a double tap. */
  tap(row: number, col: number, time: number): boolean {
    const prev = this.last
    const double = !!prev && time - prev.time <= DOUBLE_TAP_MS && prev.row === row && prev.col === col
    this.last = double ? null : { time, row, col }
    return double
  }

  reset() {
    this.last = null
  }
}

/**
 * How far the editor must scroll up to stay above the on-screen keyboard: the visual viewport shrinks when the
 * keyboard opens, and the editor's bottom (CSS px from the window top) must stay inside it with a small margin.
 */
export function keyboardOverlap(editorBottom: number, viewportTop: number, viewportHeight: number, margin = 12) {
  return Math.max(0, editorBottom + margin - (viewportTop + viewportHeight))
}
