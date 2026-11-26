import type { HeaderPicture, HeaderPictureId, PageSetup, PrintPage, Sheet } from "@cells/core"
import { headerFooterLines, type PrintContext } from "./render"

/**
 * Header and footer sections with their pictures, as Excel prints them: a section's &G puts its picture in the line,
 * after the text before it; everything on a line stands on one baseline (the picture's foot), headers hanging from
 * their margin and footers standing on theirs. Shared by printing and Page Layout view.
 */

type Loaded = { img: HTMLImageElement; ready: boolean; waiting: (() => void)[] }
const images = new Map<string, Loaded>()

/** A picture's image once it has loaded (null until then; `onReady` runs when it has). */
export function headerImage(src: string, onReady?: () => void): HTMLImageElement | null {
  let entry = images.get(src)
  if (!entry) {
    const img = new Image()
    const made: Loaded = { img, ready: false, waiting: [] }
    entry = made
    images.set(src, made)
    img.onload = () => {
      made.ready = true
      for (const run of made.waiting.splice(0)) run()
    }
    img.src = src
  }
  if (entry.ready) return entry.img
  if (onReady) entry.waiting.push(onReady)
  return null
}

/** Every picture of a page setup loaded (printing waits for them). */
export async function loadHeaderImages(setup: PageSetup): Promise<void> {
  await Promise.all(
    Object.values(setup.pictures).map(
      (p) =>
        new Promise<void>((resolve) => {
          if (headerImage(p.src, resolve)) resolve()
          // A picture that can't load prints as nothing.
          images.get(p.src)!.img.onerror = () => resolve()
        }),
    ),
  )
}

/** Format Picture's Image control as a canvas filter. */
export function pictureFilter(p: HeaderPicture, blackAndWhitePage = false): string {
  const out: string[] = []
  if (p.color !== "automatic" || blackAndWhitePage) out.push("grayscale(1)")
  if (p.color === "blackAndWhite") out.push("contrast(100)")
  if (p.brightness !== 0.5) out.push(`brightness(${(0.5 + p.brightness).toFixed(3)})`)
  if (p.contrast !== 0.5) {
    const gain = p.contrast <= 0.5 ? p.contrast * 2 : 0.5 / Math.max(0.01, 1 - p.contrast)
    out.push(`contrast(${gain.toFixed(3)})`)
  }
  return out.length ? out.join(" ") : "none"
}

export type SectionDraw = {
  /** The section's text in Excel's codes. */
  text: string
  /** Its picture, when it has one (drawn where the text has &G). */
  picture: HeaderPicture | null
  align: "left" | "center" | "right"
  /** The section's edge (left, right) or middle (center). */
  x: number
  /** Header: the top it hangs from. Footer: the baseline its last line stands on. */
  y: number
  where: "header" | "footer"
  /** The text's size, and how big a point is, in the canvas's units. */
  size: number
  pt: number
  page: PrintPage
  sheet: Sheet
  ctx: PrintContext
  /** Black and white printing. */
  mono?: boolean
  onImage?: () => void
}

type Line = { before: string; after: string; pic: HeaderPicture | null; height: number; ascent: number }

/** Where a section's &G is (not "&&G", an ampersand and a G), or -1. */
export function pictureCode(text: string): number {
  for (let i = 0; i < text.length - 1; i++) {
    if (text[i] !== "&") continue
    if (text[i + 1] === "G") return i
    i++
  }
  return -1
}

/** What a section is drawn with: a canvas, or the grid's painter (which may lack pictures and measuring in tests). */
export type SectionPainter = Pick<CanvasRenderingContext2D, "fillText" | "save" | "restore"> & {
  textAlign: CanvasTextAlign
  filter?: string
  drawImage?: CanvasRenderingContext2D["drawImage"]
  measureText?: CanvasRenderingContext2D["measureText"]
}

export function drawHeaderSection(p: SectionPainter, s: SectionDraw) {
  const measure = (t: string) => (p.measureText ? p.measureText(t).width : t.length * s.size * 0.5)
  const lineHeight = s.size * 1.2
  const show = (raw: string) => headerFooterLines(raw, s.page, s.sheet, s.ctx).join(" ")
  const lines: Line[] = s.text.split(/\r?\n/).map((raw) => {
    const at = s.picture ? pictureCode(raw) : -1
    if (at < 0) return { before: show(raw), after: "", pic: null, height: lineHeight, ascent: s.size }
    const picH = s.picture!.height * s.pt
    return {
      before: show(raw.slice(0, at)),
      after: show(raw.slice(at + 2)),
      pic: s.picture,
      height: Math.max(lineHeight, picH + s.size * 0.2),
      ascent: Math.max(s.size, picH),
    }
  })
  // Each line's baseline: headers from the top down, footers from the bottom up.
  const baselines: number[] = []
  if (s.where === "header") {
    let top = s.y
    for (const l of lines) {
      baselines.push(top + l.ascent)
      top += l.height
    }
  } else {
    // The last line stands on the footer's baseline; each line above sits its ascent plus the line below's room
    // under its baseline higher.
    baselines[lines.length - 1] = s.y
    for (let i = lines.length - 1; i > 0; i--)
      baselines[i - 1] = baselines[i] - lines[i].ascent - (lines[i - 1].height - lines[i - 1].ascent)
  }
  p.textAlign = "left"
  lines.forEach((l, i) => {
    const picW = l.pic ? l.pic.width * s.pt : 0
    const wBefore = l.before ? measure(l.before) : 0
    const wAfter = l.after ? measure(l.after) : 0
    const width = wBefore + picW + wAfter
    let x = s.align === "left" ? s.x : s.align === "center" ? s.x - width / 2 : s.x - width
    const base = baselines[i]
    if (l.before) p.fillText(l.before, x, base)
    x += wBefore
    if (l.pic) {
      const img = p.drawImage ? headerImage(l.pic.src, s.onImage) : null
      if (img && p.drawImage) {
        const c = l.pic.crop
        const nw = img.naturalWidth || 1
        const nh = img.naturalHeight || 1
        p.save()
        p.filter = pictureFilter(l.pic, s.mono)
        p.drawImage(
          img,
          c.left * nw,
          c.top * nh,
          Math.max(1, nw * (1 - c.left - c.right)),
          Math.max(1, nh * (1 - c.top - c.bottom)),
          x,
          base - l.pic.height * s.pt,
          picW,
          l.pic.height * s.pt,
        )
        p.restore()
      }
      x += picW
    }
    if (l.after) p.fillText(l.after, x, base)
  })
}

/** The picture a section prints, if its text asks for one. */
export function sectionPicture(setup: PageSetup, id: HeaderPictureId, text: string): HeaderPicture | null {
  return pictureCode(text) >= 0 ? (setup.pictures[id] ?? null) : null
}
