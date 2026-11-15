/**
 * PNG pictures of charts for the clipboard. A copy event has to fill the clipboard at once, and drawing a picture
 * takes a moment, so the chart layer draws one when a chart is selected (and again when it changes); Copy uses it if
 * it matches what the chart looks like now.
 */

const SCALE = 2
const cache = new Map<number, { key: string; url: string }>()

/** The picture key: what the chart draws and its size. */
export const pngKey = (svg: string, w: number, h: number) => `${w}x${h}:${svg}`

/** The ready PNG (a data URL) for a chart as it looks now, or null. */
export function readyPng(chartId: number, key: string): string | null {
  const hit = cache.get(chartId)
  return hit && hit.key === key ? hit.url : null
}

/** Draws the chart's PNG in the background (twice the size, for sharp pastes), unless it is ready. */
export async function preparePng(chartId: number, svg: string, w: number, h: number): Promise<void> {
  const key = pngKey(svg, w, h)
  if (readyPng(chartId, key) || !svg || typeof document === "undefined") return
  const markup = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${svg}</svg>`
  const img = new Image()
  img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(markup)
  try {
    await img.decode()
    const canvas = document.createElement("canvas")
    canvas.width = Math.round(w * SCALE)
    canvas.height = Math.round(h * SCALE)
    const p = canvas.getContext("2d")
    if (!p) return
    p.fillStyle = "#ffffff"
    p.fillRect(0, 0, canvas.width, canvas.height)
    p.drawImage(img, 0, 0, canvas.width, canvas.height)
    cache.set(chartId, { key, url: canvas.toDataURL("image/png") })
  } catch {
    // No PNG: Copy carries the SVG picture.
  }
}
