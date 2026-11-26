import type { OcrLine } from "./ocr"
import { OCR_DPI, runsFromOcr } from "./ocr"
import type { Recognizer } from "./pdf"
import type { PictureReader } from "./picture"

/**
 * Text recognition on Android: Google's ML Kit, on the device, through MainActivity's `window.CellsOcr`
 * (src-tauri/gen/android/.../AndroidOcr.kt). The PDF is handed over once and Android draws each scanned page at 300
 * dpi itself (drawing it in the web view and encoding a picture takes seconds a page); the words come back to
 * `window.__cellsOcr` in the shape Tesseract's are read in, and go to the same table finder.
 */

type Bridge = {
  available(): boolean
  recognize(id: string, jpeg: string): void
  recognizeImage?(id: string, image: string): void
  openPdf?(pdf: string): boolean
  recognizePdfPage?(id: string, index: number, scale: number): void
  closePdf?(): void
}
type Answer = { lines?: OcrLine[]; error?: string }
type Host = { CellsOcr?: Bridge; __cellsOcr?: (id: string, json: string) => void }

const host = () => globalThis as Host

/** The bridge, on Android only. */
export const androidOcr = (): Bridge | null => host().CellsOcr ?? null

const waiting = new Map<string, (answer: Answer) => void>()
let count = 0

/** Reads one picture (a JPEG, base64) with the bridge. */
export function recognizeJpeg(bridge: Bridge, jpeg: string): Promise<OcrLine[]> {
  return ask((id) => bridge.recognize(id, jpeg))
}

/** Starts a reading with `start(id)` and waits for its answer. */
function ask(start: (id: string) => void): Promise<OcrLine[]> {
  host().__cellsOcr ??= (id, json) => {
    const done = waiting.get(id)
    waiting.delete(id)
    let answer: Answer
    try {
      answer = JSON.parse(json) as Answer
    } catch {
      answer = { error: "Text recognition gave an answer Cells couldn't read." }
    }
    done?.(answer)
  }
  const id = `ocr${++count}`
  return new Promise((resolve, reject) => {
    waiting.set(id, (a) =>
      a.error || !a.lines ? reject(new Error(a.error ?? "Text recognition failed.")) : resolve(a.lines),
    )
    start(id)
  })
}

const base64 = (bytes: Uint8Array) => {
  let text = ""
  for (let i = 0; i < bytes.length; i += 0x8000) text += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return btoa(text)
}

/** The PDF the bridge holds, and when it lets it go (a minute after the last page). */
let held: Uint8Array | null = null
let letGo: ReturnType<typeof setTimeout> | undefined

/** A canvas as a JPEG in base64 (without the data: prefix). */
async function jpegOf(canvas: HTMLCanvasElement): Promise<string> {
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.92))
  if (!blob) throw new Error("The page couldn't be drawn for recognition.")
  return base64(new Uint8Array(await blob.arrayBuffer()))
}

/** Reads a scanned page with ML Kit. */
export const mlkitRecognizer: Recognizer = async (page, draw, pdf) => {
  const bridge = androidOcr()
  if (!bridge) throw new Error("Text recognition isn't available on this device.")
  const scale = OCR_DPI / 72
  if (bridge.openPdf && bridge.recognizePdfPage && bridge.closePdf) {
    clearTimeout(letGo)
    if (held !== pdf.bytes) {
      if (!bridge.openPdf?.(base64(pdf.bytes))) throw new Error("The PDF couldn't be handed to text recognition.")
      held = pdf.bytes
    }
    try {
      const lines = await ask((id) => bridge.recognizePdfPage?.(id, pdf.index, scale))
      // Android's renderer draws some PDFs blank that pdf.js draws (a picture whose colours the file gets wrong):
      // with no text found, the page is drawn here and read again.
      if (lines.length) return { ...page, runs: runsFromOcr(lines, scale) }
    } finally {
      letGo = setTimeout(() => {
        held = null
        bridge.closePdf?.()
      }, 60_000)
    }
  }
  return fromPicture(bridge, page, draw, scale)
}

/** A page drawn here, handed over as a JPEG (seconds a page in a web view, so only when Android can't draw it). */
async function fromPicture(
  bridge: Bridge,
  page: Parameters<Recognizer>[0],
  draw: Parameters<Recognizer>[1],
  scale: number,
): Promise<Awaited<ReturnType<Recognizer>>> {
  const canvas = await draw(scale)
  try {
    return { ...page, runs: runsFromOcr(await recognizeJpeg(bridge, await jpegOf(canvas)), scale) }
  } finally {
    canvas.width = 0
  }
}

/** Reads a picture (Data › From Picture) with ML Kit, which opens the file itself (a photo turned upright, HEIC too). */
export const mlkitPicture: PictureReader = async (image) => {
  const bridge = androidOcr()
  if (!bridge?.recognizeImage) throw new Error("Text recognition isn't available on this device.")
  const data = base64(new Uint8Array(await image.arrayBuffer()))
  return ask((id) => bridge.recognizeImage?.(id, data))
}
