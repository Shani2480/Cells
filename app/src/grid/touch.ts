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

/** Fling: the sheet glides on after a quick pan, slowing as a phone's lists do (iOS and Android decay alike). */
export const FLING_TIME_CONSTANT = 325
/** Below this speed (sheet px per ms) the glide stops; a slow lift doesn't glide at all. */
export const FLING_MIN_SPEED = 0.05
export const FLING_MAX_SPEED = 8
/** Only the pan's last moments count toward its speed: a pause before lifting stops the glide. */
const FLING_WINDOW_MS = 100

export type Velocity = { x: number; y: number }

/** The scroll positions of a pan over time, giving its speed when the finger lifts. */
export class FlingTracker {
  private samples: { x: number; y: number; t: number }[] = []

  add(x: number, y: number, t: number) {
    this.samples.push({ x, y, t })
    while (this.samples.length > 2 && t - this.samples[0].t > FLING_WINDOW_MS) this.samples.shift()
  }

  /** The speed at the lift (scroll px per ms), or null for a pan too slow, too short or paused before lifting. */
  release(t: number): Velocity | null {
    const samples = this.samples
    this.samples = []
    if (samples.length < 2) return null
    const last = samples[samples.length - 1]
    if (t - last.t > FLING_WINDOW_MS / 2) return null
    const first = samples[0]
    const dt = last.t - first.t
    if (dt <= 0) return null
    const clamp = (v: number) => Math.max(-FLING_MAX_SPEED, Math.min(FLING_MAX_SPEED, v))
    const v = { x: clamp((last.x - first.x) / dt), y: clamp((last.y - first.y) / dt) }
    return Math.hypot(v.x, v.y) < FLING_MIN_SPEED ? null : v
  }
}

/** One frame of a glide: how far to scroll and the speed after it (null when it has stopped). */
export function flingStep(v: Velocity, dt: number): { dx: number; dy: number; next: Velocity | null } {
  const decay = Math.exp(-dt / FLING_TIME_CONSTANT)
  // The distance covered while slowing from v to v·decay over dt.
  const travel = FLING_TIME_CONSTANT * (1 - decay)
  const next = { x: v.x * decay, y: v.y * decay }
  return {
    dx: v.x * travel,
    dy: v.y * travel,
    next: Math.hypot(next.x, next.y) < FLING_MIN_SPEED ? null : next,
  }
}
