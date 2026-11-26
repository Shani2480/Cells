import type { PdfRun } from "@cells/io"
import type { Recognizer } from "./pdf"
import type { PictureReader } from "./picture"

/**
 * Text recognition for scanned PDF pages (opt-in: File › Options › AI): the page drawn at 300 dpi and read by
 * Tesseract in a worker, on this computer, with the English model that ships with Cells (`vite.config.ts` copies the
 * files). Its words go to the same table finder as a PDF's own text.
 */

/** Tesseract's result, as far as it's used here: blocks of paragraphs of lines of words, boxes in pixels. */
export type OcrBox = { x0: number; y0: number; x1: number; y1: number }
export type OcrWord = { text: string; confidence: number; bbox: OcrBox }
/** `angle`: how far the line is turned, in degrees clockwise (ML Kit gives it; Tesseract's pages come straight). */
export type OcrLine = { bbox: OcrBox; words: OcrWord[]; angle?: number }
export type OcrPage = { blocks: { paragraphs: { lines: OcrLine[] }[] }[] | null }

/** The dpi pages are drawn at for recognition (Tesseract reads best at 300). */
export const OCR_DPI = 300

/** Below this, a chunk is read again on its own, and a word still below it is marked for a person to check. */
const SURE = 60
/** Below this a word is left out: what's read is more likely a speck than text. */
const NOISE = 10

const isRule = (t: string) => /^([|_¦]+|[—–=~-]{2,})$/.test(t)

/** The part of a Tesseract worker this uses (the app's runs in the browser, the tests' in Node). */
export type OcrEngine = {
  recognize(
    image: unknown,
    options?: { rotateAuto?: boolean; rectangle?: { left: number; top: number; width: number; height: number } },
    output?: object,
  ): Promise<{ data: OcrPage }>
  setParameters(params: Record<string, string>): Promise<unknown>
}

const linesIn = (page: OcrPage) => (page.blocks ?? []).flatMap((b) => b.paragraphs.flatMap((p) => p.lines))
const mean = (words: OcrWord[]) => words.reduce((s, w) => s + w.confidence, 0) / Math.max(1, words.length)

/**
 * Reads a page: first as sparse text (every chunk of words found where it is, rather than the page read as one
 * block, which runs a table's rows together), straightened first, since a scan is rarely square. Then each chunk
 * Tesseract was unsure of is read again on its own as one line, which is often right where the first reading
 * wasn't ("2 €18.50" for "25 €18.50"); the better reading is kept.
 */
export async function recognizePage(engine: OcrEngine, image: unknown): Promise<OcrLine[]> {
  await engine.setParameters({ tessedit_pageseg_mode: "11" })
  const first = await engine.recognize(image, { rotateAuto: true }, { blocks: true, text: false })
  const lines = linesIn(first.data)
  const doubtful = lines.filter((l) => l.words.some((w) => w.confidence < SURE))
  if (!doubtful.length) return lines
  await engine.setParameters({ tessedit_pageseg_mode: "7" })
  for (const line of doubtful) {
    const pad = Math.round((line.bbox.y1 - line.bbox.y0) * 0.3)
    const rectangle = {
      left: Math.max(0, line.bbox.x0 - pad),
      top: Math.max(0, line.bbox.y0 - pad),
      width: line.bbox.x1 - line.bbox.x0 + 2 * pad,
      height: line.bbox.y1 - line.bbox.y0 + 2 * pad,
    }
    const again = linesIn((await engine.recognize(image, { rectangle }, { blocks: true, text: false })).data)
    const words = again.flatMap((l) => l.words).filter((w) => w.text.trim())
    if (words.length && mean(words) > mean(line.words)) line.words = words
  }
  return lines
}

const middle = (values: number[]) => {
  const s = [...values].sort((a, b) => a - b)
  return s.length ? s[Math.floor(s.length / 2)] : 0
}

/**
 * How far a page is turned, in degrees clockwise: the middle of its lines' angles (lines of two or more words). 0 when
 * the recognizer gives none (Tesseract straightens the page itself).
 */
export function pageAngle(lines: OcrLine[]): number {
  const given = lines.flatMap((l) => (typeof l.angle === "number" && l.words.length >= 2 ? [l.angle] : []))
  return given.length >= 3 ? middle(given) : 0
}

/**
 * The lines turned back by the page's angle (about its top left), so a column's edge runs straight down the page and
 * the table finder sees its words line up. A scan is rarely square; over a page, 0.4° moves a column by about a word
 * gap. Pages turned less than 0.05° are left as they are.
 */
export function straighten(lines: OcrLine[]): OcrLine[] {
  const angle = pageAngle(lines)
  if (Math.abs(angle) < 0.05) return lines
  const r = (-angle * Math.PI) / 180
  const cos = Math.cos(r)
  const sin = Math.sin(r)
  // A box turned, as the box around its turned corners' middle, kept its size (the corners would widen it).
  const turn = (b: OcrBox): OcrBox => {
    const cx = (b.x0 + b.x1) / 2
    const cy = (b.y0 + b.y1) / 2
    const x = cx * cos - cy * sin
    const y = cx * sin + cy * cos
    const hw = (b.x1 - b.x0) / 2
    const hh = (b.y1 - b.y0) / 2
    return { x0: x - hw, y0: y - hh, x1: x + hw, y1: y + hh }
  }
  return lines.map((l) => ({
    ...l,
    angle: 0,
    bbox: turn(l.bbox),
    words: l.words.map((w) => ({ ...w, bbox: turn(w.bbox) })),
  }))
}

/**
 * A recognized page's text as runs in points: a run for each chunk of text (words close together), with each word's
 * place, sized by the chunk's height. Rule marks (| _ ¦ or a run of dashes, from a table's lines) and specks are left
 * out; words recognition wasn't sure of are marked.
 */
export function runsFromOcr(lines: OcrLine[], scale: number): PdfRun[] {
  const runs: PdfRun[] = []
  for (const line of straighten(lines)) {
    const words = line.words
      .map((w) => ({ ...w, text: w.text.trim() }))
      .filter((w) => w.text && !isRule(w.text) && (w.confidence >= NOISE || /[\p{L}\p{N}]{2,}/u.test(w.text)))
    if (!words.length) continue
    runs.push({
      text: words.map((w) => w.text).join(" "),
      x0: Math.min(...words.map((w) => w.bbox.x0)) / scale,
      x1: Math.max(...words.map((w) => w.bbox.x1)) / scale,
      y0: line.bbox.y0 / scale,
      y1: line.bbox.y1 / scale,
      size: (line.bbox.y1 - line.bbox.y0) / scale,
      words: words.map((w) => ({
        text: w.text,
        x0: w.bbox.x0 / scale,
        x1: w.bbox.x1 / scale,
        ...(w.confidence < SURE ? { unsure: true } : {}),
      })),
    })
  }
  return runs
}

type TesseractWorker = OcrEngine & { terminate(): Promise<unknown> }

let worker: Promise<TesseractWorker> | null = null
let idle: ReturnType<typeof setTimeout> | undefined

/** The worker, started on first use and stopped after a minute without work (it holds the model in memory). */
function tesseract(): Promise<TesseractWorker> {
  clearTimeout(idle)
  worker ??= import("tesseract.js").then(({ createWorker, OEM }) => {
    const at = (p: string) => new URL(p, window.location.href).href
    return createWorker("eng", OEM.LSTM_ONLY, {
      workerPath: at("/ocr/worker.min.js"),
      corePath: at("/ocr/core"),
      langPath: at("/ocr/lang"),
      workerBlobURL: false,
      cacheMethod: "none",
      gzip: true,
    }) as unknown as Promise<TesseractWorker>
  })
  const current = worker
  current.catch(() => {
    if (worker === current) worker = null
  })
  return current
}

function release() {
  clearTimeout(idle)
  idle = setTimeout(() => {
    const w = worker
    worker = null
    void w?.then((t) => t.terminate())
  }, 60_000)
}

/** Reads a scanned page with Tesseract. */
export const tesseractRecognizer: Recognizer = async (page, draw) => {
  const scale = OCR_DPI / 72
  const canvas = await draw(scale)
  try {
    const t = await tesseract()
    return { ...page, runs: runsFromOcr(await recognizePage(t, canvas), scale) }
  } finally {
    canvas.width = 0
    release()
  }
}

/**
 * A picture drawn for Tesseract, upright (createImageBitmap turns a photo as its camera noted). A screenshot is
 * enlarged (Tesseract reads text about 30 pixels high best, a screen's is 12 to 16) and a large photo made smaller.
 */
async function pictureCanvas(image: Blob): Promise<HTMLCanvasElement> {
  const bitmap = await createImageBitmap(image).catch(() => {
    throw new Error("Cells can't open this kind of picture. Try a PNG or JPEG.")
  })
  const long = Math.max(bitmap.width, bitmap.height)
  const scale = long < 2000 ? Math.min(3, 2400 / long) : long > 4200 ? 4200 / long : 1
  const canvas = document.createElement("canvas")
  canvas.width = Math.round(bitmap.width * scale)
  canvas.height = Math.round(bitmap.height * scale)
  const ctx = canvas.getContext("2d")
  if (!ctx) throw new Error("The picture can't be drawn here.")
  ctx.fillStyle = "#fff"
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.imageSmoothingQuality = "high"
  ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height)
  bitmap.close()
  return canvas
}

/** Reads a picture (Data › From Picture) with Tesseract. */
export const tesseractPicture: PictureReader = async (image) => {
  const canvas = await pictureCanvas(image)
  try {
    return await recognizePage(await tesseract(), canvas)
  } finally {
    canvas.width = 0
    release()
  }
}
