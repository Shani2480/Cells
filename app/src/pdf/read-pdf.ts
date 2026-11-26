import type { PdfPageText, PdfRule, PdfRun } from "@cells/io"

/**
 * A PDF's text and ruling lines, page by page, read with pdf.js for From PDF (`@cells/io`'s pdfTables finds the
 * tables in them). Positions are in points from the page's top left, as the page shows (rotated pages turned).
 * A page with (almost) no text but a picture is a scan: its text has to be recognized first (`ocr.ts`).
 */

export type PdfPage = PdfPageText & { scanned: boolean }

/** The part of pdf.js used here (the app passes the browser build, tests the Node one). */
export type PdfJs = {
  getDocument(src: { data: Uint8Array; isEvalSupported?: boolean; disableFontFace?: boolean }): {
    promise: Promise<PdfDocument>
    destroy(): Promise<void>
  }
  OPS: Record<string, number>
  Util: { transform(a: number[], b: number[]): number[] }
}

export type PdfDocument = {
  numPages: number
  getPage(n: number): Promise<PdfPageProxy>
}

export type PdfPageProxy = {
  getViewport(o: { scale: number }): { width: number; height: number; transform: number[] }
  getTextContent(): Promise<{ items: unknown[] }>
  getOperatorList(): Promise<{ fnArray: number[]; argsArray: unknown[] }>
  render(o: { canvasContext: unknown; viewport: unknown; canvas?: unknown }): { promise: Promise<void> }
}

type TextItem = { str: string; transform: number[]; width: number; height: number }

const isText = (i: unknown): i is TextItem =>
  typeof i === "object" && i !== null && "str" in i && "transform" in i && "width" in i

/** The page's text runs, left to right text only (turned and vertical text is left out). */
function runsOf(items: unknown[], view: number[], lib: PdfJs): PdfRun[] {
  const runs: PdfRun[] = []
  for (const item of items) {
    if (!isText(item) || !item.str.trim()) continue
    const t = lib.Util.transform(view, item.transform)
    // Horizontal text reads as [a, ~0, ~0, d] once the page is turned upright.
    if (Math.abs(t[1]) > 0.01 * Math.abs(t[0]) || Math.abs(t[2]) > 0.01 * Math.abs(t[3]) || t[0] <= 0) continue
    const size = Math.abs(t[3]) || Math.abs(t[0])
    const x0 = t[4]
    const baseline = t[5]
    runs.push({ text: item.str, x0, x1: x0 + item.width, y0: baseline - 0.8 * size, y1: baseline + 0.2 * size, size })
  }
  return runs
}

type Matrix = number[]
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]
const apply = (m: Matrix, x: number, y: number): [number, number] => [
  m[0] * x + m[2] * y + m[4],
  m[1] * x + m[3] * y + m[5],
]

/** The ruling lines: stroked straight lines (a rectangle's sides too) and thin filled rectangles. */
function rulesOf(ops: { fnArray: number[]; argsArray: unknown[] }, view: number[], lib: PdfJs): PdfRule[] {
  const O = lib.OPS
  const stroke = new Set([
    O.stroke,
    O.closeStroke,
    O.fillStroke,
    O.eoFillStroke,
    O.closeFillStroke,
    O.closeEOFillStroke,
  ])
  const fill = new Set([O.fill, O.eoFill, O.fillStroke, O.eoFillStroke, O.closeFillStroke, O.closeEOFillStroke])
  const rules: PdfRule[] = []
  let ctm: Matrix = IDENTITY
  const stack: Matrix[] = []
  const toPage = (x: number, y: number) => {
    const [px, py] = apply(ctm, x, y)
    return apply(view, px, py)
  }
  const straight = (a: [number, number], b: [number, number]) => {
    if (Math.abs(a[0] - b[0]) < 0.5 || Math.abs(a[1] - b[1]) < 0.5)
      rules.push({
        x0: Math.min(a[0], b[0]),
        y0: Math.min(a[1], b[1]),
        x1: Math.max(a[0], b[0]),
        y1: Math.max(a[1], b[1]),
      })
  }
  ops.fnArray.forEach((fn, k) => {
    const args = ops.argsArray[k] as unknown[]
    if (fn === O.save) stack.push(ctm)
    else if (fn === O.restore) ctm = stack.pop() ?? IDENTITY
    else if (fn === O.transform) ctm = lib.Util.transform(ctm, args as number[])
    else if (fn === O.paintFormXObjectBegin) {
      stack.push(ctm)
      if (Array.isArray(args?.[0])) ctm = lib.Util.transform(ctm, args[0] as number[])
    } else if (fn === O.paintFormXObjectEnd) ctm = stack.pop() ?? IDENTITY
    else if (fn === O.constructPath) {
      const paint = args[0] as number
      if (!stroke.has(paint) && !fill.has(paint)) return
      const data = (Array.isArray(args[1]) ? args[1][0] : args[1]) as ArrayLike<number> | undefined
      if (!data) return
      // Subpaths: 0 move, 1 line, 2 curve, 3 quadratic curve, 4 close.
      let sub: [number, number][] = []
      let curved = false
      const end = (closed: boolean) => {
        if (sub.length >= 2) {
          if (stroke.has(paint)) {
            for (let i = 1; i < sub.length; i++) straight(sub[i - 1], sub[i])
            if (closed) straight(sub[sub.length - 1], sub[0])
          }
          if (fill.has(paint) && !curved && sub.length >= 4 && sub.length <= 5) {
            const xs = sub.map((p) => p[0])
            const ys = sub.map((p) => p[1])
            const r = { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) }
            if (r.x1 - r.x0 <= 3 || r.y1 - r.y0 <= 3) rules.push(r)
          }
        }
        sub = []
        curved = false
      }
      for (let i = 0; i < data.length;) {
        const op = data[i++]
        if (op === 0) {
          end(false)
          sub.push(toPage(data[i], data[i + 1]))
          i += 2
        } else if (op === 1) {
          sub.push(toPage(data[i], data[i + 1]))
          i += 2
        } else if (op === 2) {
          curved = true
          sub.push(toPage(data[i + 4], data[i + 5]))
          i += 6
        } else if (op === 3) {
          curved = true
          sub.push(toPage(data[i + 2], data[i + 3]))
          i += 4
        } else if (op === 4) end(true)
        else break
      }
      end(false)
    }
  })
  return rules
}

const IMAGE_OPS = ["paintImageXObject", "paintInlineImageXObject", "paintImageMaskXObject", "paintImageXObjectRepeat"]

/** Reads a PDF's pages. Rejects with pdf.js's error for a file that isn't a PDF (or needs a password). */
export async function readPdf(
  bytes: Uint8Array,
  lib: PdfJs,
): Promise<{ pages: PdfPage[]; doc: PdfDocument; close: () => Promise<void> }> {
  // pdf.js takes the buffer over, so it gets a copy.
  const task = lib.getDocument({ data: bytes.slice(), isEvalSupported: false, disableFontFace: true })
  const doc = await task.promise
  const pages: PdfPage[] = []
  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n)
    const viewport = page.getViewport({ scale: 1 })
    const [text, ops] = await Promise.all([page.getTextContent(), page.getOperatorList()])
    const runs = runsOf(text.items, viewport.transform, lib)
    const images = new Set(IMAGE_OPS.map((name) => lib.OPS[name]))
    const pictured = ops.fnArray.some((f) => images.has(f))
    const chars = runs.reduce((s, r) => s + r.text.trim().length, 0)
    pages.push({
      width: viewport.width,
      height: viewport.height,
      runs,
      rules: rulesOf(ops, viewport.transform, lib),
      scanned: pictured && chars < 20,
    })
  }
  return { pages, doc, close: () => task.destroy() }
}
