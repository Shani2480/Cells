import type { Pagination, PrintPage, Sheet } from "@cells/core"
import { pdfFromImages, type PdfPage } from "@cells/io"
import type { SheetController } from "@/sheet/controller"
import { drawPage, type PrintContext } from "./render"

/** File › Print › Settings: what prints. */
export type PrintWhat = "active" | "workbook" | "selection"

/** One page of a print job, with the sheet and pagination it belongs to. */
export type JobPage = { sheet: Sheet; pagination: Pagination; page: PrintPage }

/**
 * The pages a print job prints: the active sheet, every visible sheet, or the selection. Automatic page numbers run on
 * from sheet to sheet, as in one Excel print job.
 */
export function printJob(c: SheetController, what: PrintWhat, ignorePrintArea = false): JobPage[] {
  const out: JobPage[] = []
  const sheets = what === "workbook" ? c.wb.sheets.filter((s) => !s.hidden) : [c.sheet]
  let next = 1
  for (const sheet of sheets) {
    const ranges = what === "selection" ? c.selection.ranges.map((r) => ({ ...r })) : undefined
    const pagination = c.print.pages(sheet, { ranges, ignorePrintArea, firstNumber: next })
    for (const page of pagination.pages) out.push({ sheet, pagination, page })
    const last = pagination.pages.at(-1)
    if (last) next = last.number + 1
  }
  return out
}

/** A page drawn on a new canvas at `pxPerPt` pixels per point. */
export async function pageCanvas(
  c: SheetController,
  job: JobPage,
  ctx: PrintContext,
  pxPerPt: number,
): Promise<HTMLCanvasElement> {
  const canvas = document.createElement("canvas")
  canvas.width = Math.round(job.pagination.paper.width * pxPerPt)
  canvas.height = Math.round(job.pagination.paper.height * pxPerPt)
  await drawPage(canvas, c, job.sheet, job.pagination, job.page, ctx, pxPerPt)
  return canvas
}

/** Print and PDF resolution: 200 dots per inch keeps 11 pt text sharp at a fair file size. */
export const PRINT_DPI = 200

const blobOf = (canvas: HTMLCanvasElement, type: string, quality?: number) =>
  new Promise<Blob>((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("The page couldn't be drawn."))), type, quality),
  )

/** The job as a PDF: each page drawn at print resolution. */
export async function jobPdf(
  c: SheetController,
  pages: readonly JobPage[],
  ctx: PrintContext,
  progress?: (done: number) => void,
): Promise<Uint8Array> {
  const pxPerPt = PRINT_DPI / 72
  const out: PdfPage[] = []
  for (const [i, job] of pages.entries()) {
    const canvas = await pageCanvas(c, job, ctx, pxPerPt)
    const jpeg = new Uint8Array(await (await blobOf(canvas, "image/jpeg", 0.92)).arrayBuffer())
    out.push({
      width: job.pagination.paper.width,
      height: job.pagination.paper.height,
      jpeg,
      pixelWidth: canvas.width,
      pixelHeight: canvas.height,
    })
    progress?.((i + 1) / pages.length)
    // Let the window draw the progress between pages.
    await new Promise((r) => setTimeout(r, 0))
  }
  return pdfFromImages(out, { title: ctx.fileName.replace(/\.[^.]+$/, ""), date: ctx.now })
}

/**
 * Print: the pages as pictures in a frame of their own, sized to their paper, handed to the system's print dialog
 * (which also offers printing to a PDF file).
 */
export async function printJobPages(c: SheetController, pages: readonly JobPage[], ctx: PrintContext) {
  if (!pages.length) return
  const pxPerPt = PRINT_DPI / 72
  const urls: string[] = []
  for (const job of pages)
    urls.push(URL.createObjectURL(await blobOf(await pageCanvas(c, job, ctx, pxPerPt), "image/png")))
  const frame = document.createElement("iframe")
  frame.setAttribute("aria-hidden", "true")
  frame.style.cssText = "position:fixed;width:0;height:0;border:0;left:-10000px;top:0"
  document.body.append(frame)
  const doc = frame.contentDocument!
  const first = pages[0].pagination.paper
  // One @page size per job (the first page's); pages of other sizes are scaled into it by the printer.
  doc.open()
  doc.write(
    `<!doctype html><html><head><title>${escapeHtml(ctx.fileName)}</title><style>` +
      `@page{size:${first.width}pt ${first.height}pt;margin:0}html,body{margin:0;padding:0}` +
      `img{display:block;width:${first.width}pt;height:${first.height}pt;page-break-after:always;break-after:page}` +
      `img:last-child{page-break-after:auto;break-after:auto}</style></head><body>` +
      urls.map((u) => `<img src="${u}" alt="">`).join("") +
      `</body></html>`,
  )
  doc.close()
  await Promise.all([...doc.images].map((img) => img.decode().catch(() => undefined)))
  const done = () => {
    frame.remove()
    for (const u of urls) URL.revokeObjectURL(u)
  }
  frame.contentWindow!.addEventListener("afterprint", () => setTimeout(done, 0), { once: true })
  frame.contentWindow!.focus()
  frame.contentWindow!.print()
  // Some webviews never send afterprint: clean up after a while anyway.
  setTimeout(done, 60_000)
}

const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (ch) => `&#${ch.charCodeAt(0)};`)
