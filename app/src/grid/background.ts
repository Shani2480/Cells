import type { GridLayout } from "./layout"
import type { SheetController } from "@/sheet/controller"

/**
 * Page Layout › Background: the sheet's picture tiled behind the cells, from A1, at its own size in sheet pixels (so
 * it scales with the zoom, as in Excel). Fills, gridlines, text and objects draw over it. Excel shows it in Normal and
 * Page Break Preview, not in Page Layout view or on paper.
 */

/** How much of the dark sheet colour covers the picture in dark mode. */
const DIM = 0.7

type Loaded = { img: HTMLImageElement; pattern: CanvasPattern | null; ctx: unknown }

const cache = new Map<string, Loaded | "loading" | "failed">()
const waiting = new Set<() => void>()

/** Calls `redraw` once a background picture that was still loading has arrived; returns the unsubscribe. */
export function onBackgroundLoaded(redraw: () => void): () => void {
  waiting.add(redraw)
  return () => waiting.delete(redraw)
}

function loaded(src: string): Loaded | null {
  const entry = cache.get(src)
  if (entry && entry !== "loading" && entry !== "failed") return entry
  if (entry || typeof Image === "undefined") return null
  cache.set(src, "loading")
  const img = new Image()
  img.onload = () => {
    // One picture at a time: a sheet has one background, and the data URLs are large.
    for (const [key, value] of cache) if (key !== src && value !== "loading") cache.delete(key)
    cache.set(src, { img, pattern: null, ctx: null })
    for (const redraw of [...waiting]) redraw()
  }
  img.onerror = () => cache.set(src, "failed")
  img.src = src
  return null
}

/** The painter calls drawing a background needs. */
export type BackgroundPainter = {
  save(): void
  restore(): void
  beginPath(): void
  rect(x: number, y: number, w: number, h: number): void
  clip(): void
  translate(x: number, y: number): void
  fillRect(x: number, y: number, w: number, h: number): void
  createPattern?(image: CanvasImageSource, repetition: string | null): CanvasPattern | null
  fillStyle: CanvasRenderingContext2D["fillStyle"]
  globalAlpha: number
}

/**
 * Tiles the sheet's background over one pane (before its cells). In dark mode (`dim`: the sheet's dark colour) a wash
 * of that colour goes over it, so light text stays readable on a light picture.
 */
export function drawBackground(p: BackgroundPainter, controller: SheetController, pane: GridLayout, dim?: string) {
  const src = controller.sheet.background
  if (!src || pane.pages || !pane.rows.length || !pane.cols.length || !p.createPattern) return
  const entry = loaded(src)
  if (!entry || !entry.img.naturalWidth) return
  if (!entry.pattern || entry.ctx !== p) {
    entry.pattern = p.createPattern(entry.img, "repeat")
    entry.ctx = p
  }
  if (!entry.pattern) return
  const sheet = controller.sheet
  // A1's corner on the canvas: the pane's first line less its offset into the sheet.
  const x = pane.cols[0].start - sheet.cols.offset(pane.cols[0].index)
  const y = pane.rows[0].start - sheet.rows.offset(pane.rows[0].index)
  const { clip } = pane
  p.save()
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  p.clip()
  p.translate(x, y)
  p.fillStyle = entry.pattern
  p.fillRect(clip.x0 - x, clip.y0 - y, clip.x1 - clip.x0, clip.y1 - clip.y0)
  if (dim) {
    p.globalAlpha = DIM
    p.fillStyle = dim
    p.fillRect(clip.x0 - x, clip.y0 - y, clip.x1 - clip.x0, clip.y1 - clip.y0)
  }
  p.restore()
}
