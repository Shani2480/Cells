/**
 * The rows a mouse-wheel turn scrolls (three a notch, negative going up), or null for a touchpad's smooth scroll.
 * Firefox gives a notch as lines; Chrome and Safari give pixels, with wheelDeltaY in steps of 120.
 */
export function wheelRows(event: Pick<WheelEvent, "deltaMode" | "deltaX" | "deltaY"> & { wheelDeltaY?: number }) {
  if (event.deltaX !== 0 || event.deltaY === 0) return null
  if (event.deltaMode === 1) return Math.round(event.deltaY)
  const notches = event.wheelDeltaY
  if (event.deltaMode === 0 && notches && notches % 120 === 0) return (-notches / 120) * 3
  return null
}
