import type { PdfText } from "@cells/io"

/** A font's size in canvas units, from its CSS shorthand ("bold 14.67px Aptos", "11pt Calibri"). */
function fontSize(font: string): number {
  const m = /(\d+(?:\.\d+)?)(px|pt)\b/.exec(font)
  if (!m) return 10
  return m[2] === "pt" ? (+m[1] * 4) / 3 : +m[1]
}

/** How far below the textBaseline line the alphabetic baseline lies, as a share of the font size, when unmeasured. */
const BASELINE_DROP: Partial<Record<CanvasTextBaseline, number>> = {
  top: 0.8,
  hanging: 0.65,
  middle: 0.3,
  bottom: -0.2,
  ideographic: -0.2,
}

/**
 * Records the text drawn on a canvas, for an exported PDF's invisible text layer: each fillText while recording, as
 * where its baseline starts on the page, its size, width and direction, in points (`pxPerPt` canvas pixels a point).
 * Text drawn wholly off the page is left out. Stop gives what was drawn and puts fillText back.
 */
export function recordText(
  p: CanvasRenderingContext2D,
  pxPerPt: number,
  page: { width: number; height: number },
): () => PdfText[] {
  const out: PdfText[] = []
  const fillText = p.fillText
  const record = (text: string, x: number, y: number, maxWidth?: number) => {
    if (!text.trim()) return
    const m = p.getTransform()
    const metrics = p.measureText(text)
    const width = maxWidth !== undefined ? Math.min(metrics.width, maxWidth) : metrics.width
    const size = fontSize(p.font)
    const rtl = p.direction === "rtl"
    const align = p.textAlign
    const dx =
      align === "center"
        ? -width / 2
        : align === "right" || (align === "end" && !rtl) || (align === "start" && rtl)
          ? -width
          : 0
    // The alphabetic baseline: measured against the font's box where the canvas can, estimated otherwise.
    let baseline = y
    if (p.textBaseline !== "alphabetic") {
      const ascent = metrics.fontBoundingBoxAscent
      if (typeof ascent === "number" && Number.isFinite(ascent)) {
        const before = p.textBaseline
        p.textBaseline = "alphabetic"
        baseline = y - ascent + p.measureText(text).fontBoundingBoxAscent
        p.textBaseline = before
      } else baseline = y + (BASELINE_DROP[p.textBaseline] ?? 0) * size
    }
    const at = m.transformPoint({ x: x + dx, y: baseline })
    const across = Math.hypot(m.a, m.b)
    const down = Math.hypot(m.c, m.d)
    if (!across || !down) return
    const piece: PdfText = {
      text,
      x: at.x / pxPerPt,
      y: at.y / pxPerPt,
      size: (size * down) / pxPerPt,
      width: (width * across) / pxPerPt,
    }
    const angle = (Math.atan2(m.b, m.a) * 180) / Math.PI
    if (Math.abs(angle) > 0.01) piece.angle = Math.round(angle * 1000) / 1000
    const reach = piece.width + piece.size
    if (piece.x > page.width + reach || piece.x < -reach || piece.y > page.height + reach || piece.y < -reach) return
    out.push(piece)
  }
  p.fillText = function (text: string, x: number, y: number, maxWidth?: number) {
    record(text, x, y, maxWidth)
    if (maxWidth === undefined) fillText.call(p, text, x, y)
    else fillText.call(p, text, x, y, maxWidth)
  }
  return () => {
    delete (p as { fillText?: unknown }).fillText
    if (p.fillText !== fillText) p.fillText = fillText
    return out
  }
}
