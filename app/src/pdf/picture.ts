import { pdfTables, type PdfItem } from "@cells/io"
import { runsFromOcr, type OcrLine } from "./ocr"

/**
 * Data › From Picture: a photo or screenshot of a table read into cells, through the same text recognition and table
 * finder as a scanned PDF. The recognizer gives the picture's lines in pixels; they're sized so the text is about
 * 12 points high (the finder's tolerances are in points, and a picture has no page size of its own).
 */

/** Reads a picture's text: its lines, with boxes in the picture's pixels (turned upright). */
export type PictureReader = (image: Blob) => Promise<OcrLine[]>

/** A grid read from a picture, with the cells recognition wasn't sure of as [row, column]. */
export type PictureGrid = { rows: string[][]; unsure: [number, number][] }

/** What a picture holds: its table, and everything in it when that's more than the table (a title, notes). */
export type PictureRead = { table: PictureGrid; whole: PictureGrid | null }

const TEXT_POINTS = 12

const middle = (values: number[]) => {
  const s = [...values].sort((a, b) => a - b)
  return s.length ? s[Math.floor(s.length / 2)] : 0
}

const grid = (item: PdfItem): PictureGrid => ({ rows: item.rows, unsure: item.unsure ?? [] })
const cellCount = (item: PdfItem) => item.rows.length * (item.rows[0]?.length ?? 0)

/** The table in a picture's lines (the largest one found), or every line in rows and columns when none is. */
export function pictureTables(lines: OcrLine[]): PictureRead {
  const heights = lines.filter((l) => l.words.length).map((l) => l.bbox.y1 - l.bbox.y0)
  const scale = Math.max(0.05, middle(heights) / TEXT_POINTS)
  const runs = runsFromOcr(lines, scale)
  const width = Math.max(1, ...runs.map((r) => r.x1)) + 10
  const height = Math.max(1, ...runs.map((r) => r.y1)) + 10
  const items = pdfTables([{ width, height, runs, rules: [] }])
  const page = items.find((i) => i.kind === "Page")
  const tables = items.filter((i) => i.kind === "Table" && i.rows.length >= 2)
  const best = tables.reduce<PdfItem | null>((a, t) => (!a || cellCount(t) > cellCount(a) ? t : a), null)
  const whole = page && page.rows.length ? grid(page) : { rows: [], unsure: [] }
  if (!best) return { table: whole, whole: null }
  const same = page && page.rows.length === best.rows.length && page.rows[0]?.length === best.rows[0]?.length
  return { table: grid(best), whole: same ? null : whole }
}

let clipboardImage: (() => Promise<Blob | null>) | null = null

/** How a picture is read from the clipboard, where the platform can (null: Picture From Clipboard is off). */
export function setClipboardImageReader(read: (() => Promise<Blob | null>) | null) {
  clipboardImage = read
}

export const canReadClipboardImage = () => clipboardImage !== null

/** The picture on the clipboard, or null when there's none. */
export const readClipboardImage = () => (clipboardImage ? clipboardImage() : Promise.resolve(null))
