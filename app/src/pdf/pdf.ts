import { pdfTables, type PdfItem } from "@cells/io"
import { readPdf, type PdfJs, type PdfPage } from "./read-pdf"

/**
 * From PDF's reading: pdf.js, loaded the first time a PDF is opened (it's large), then the tables found on its pages.
 * Scanned pages have no text of their own; `recognize` (text recognition, when it's turned on) reads them first.
 */

let lib: Promise<PdfJs> | null = null

function pdfjs(): Promise<PdfJs> {
  lib ??= Promise.all([
    import("pdfjs-dist/legacy/build/pdf.mjs"),
    import("pdfjs-dist/legacy/build/pdf.worker.min.mjs?url"),
  ]).then(([m, worker]) => {
    // Without workers (tests in Node), pdf.js runs its worker's code in the page.
    if (typeof Worker !== "undefined") m.GlobalWorkerOptions.workerSrc = worker.default
    return m as unknown as PdfJs
  })
  return lib
}

/** Reads a scanned page's text: the page drawn as a picture, its words found in it. */
export type Recognizer = (page: PdfPage, draw: (scale: number) => Promise<HTMLCanvasElement>) => Promise<PdfPage>

export type PdfRead = {
  items: PdfItem[]
  /** Pages (from 1) that are scans, and weren't read because text recognition is off. */
  unread: number[]
}

export class PdfError extends Error {}

/** A PDF's tables and pages. Rejects with a PdfError a person can read for a file pdf.js can't open. */
export async function pdfItems(
  bytes: Uint8Array,
  recognize: Recognizer | null,
  progress?: (page: number, of: number) => void,
): Promise<PdfRead> {
  const m = await pdfjs()
  let read: Awaited<ReturnType<typeof readPdf>>
  try {
    read = await readPdf(bytes, m)
  } catch (e) {
    const name = (e as { name?: string }).name
    if (name === "PasswordException") throw new PdfError("The PDF is protected with a password.")
    throw new PdfError("The file isn't a PDF Cells can read.")
  }
  const unread: number[] = []
  const pages: PdfPage[] = []
  const scans = read.pages.filter((p) => p.scanned).length
  let done = 0
  try {
    for (const [i, page] of read.pages.entries()) {
      if (!page.scanned) pages.push(page)
      else if (!recognize) {
        unread.push(i + 1)
        pages.push(page)
      } else {
        progress?.(++done, scans)
        const proxy = await read.doc.getPage(i + 1)
        pages.push(
          await recognize(page, async (scale) => {
            const viewport = proxy.getViewport({ scale })
            const canvas = document.createElement("canvas")
            canvas.width = Math.ceil(viewport.width)
            canvas.height = Math.ceil(viewport.height)
            const ctx = canvas.getContext("2d")
            if (!ctx) throw new PdfError("The page can't be drawn here.")
            await proxy.render({ canvasContext: ctx, viewport, canvas }).promise
            return canvas
          }),
        )
      }
    }
  } finally {
    void read.close()
  }
  return { items: pdfTables(pages), unread }
}
