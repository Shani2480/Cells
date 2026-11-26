import { contentForInput, regional, selectRange, type CellWrite, type Regional } from "@cells/core"
import { pictureTables, type PictureRead, type PictureReader } from "@/pdf/picture"
import type { SheetController } from "./controller"
import { message, recognitionSwitch } from "./get-data"
import { dateFrom, numberFrom } from "./power-query"

/**
 * Data › From Picture, as Excel's: a picture of a table (a photo, a screenshot) read with text recognition on this
 * device, shown for review beside the picture with the cells recognition wasn't sure of marked, then put into the
 * sheet at the active cell as if typed there, in one step to undo.
 */

let reader: PictureReader | null = null
/** Text recognition for pictures, while it's turned on (null: off). */
export function setPictureReader(r: PictureReader | null) {
  reader = r
}

const NUMERIC_DATE = /^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$/

/**
 * The order a column's dates are written in, when it isn't the region's: a picture of a UK statement read on a US
 * computer has 13/01/2026 further down, so every date in it is day first. Null when the region's order reads them all
 * (or the column isn't dates).
 */
export function dateOrderOf(texts: string[], r: Regional = regional()): Regional | null {
  const filled = texts.map((t) => t.trim()).filter(Boolean)
  const dates = filled.filter((t) => NUMERIC_DATE.test(t))
  if (!dates.length || dates.length < 0.6 * filled.length) return null
  if (dates.every((t) => dateFrom(t, r) !== null)) return null
  for (const dateOrder of ["dmy", "mdy", "ymd"] as const) {
    const other = { ...r, dateOrder }
    if (dateOrder !== r.dateOrder && dates.every((t) => dateFrom(t, other) !== null)) return other
  }
  return null
}

/** A date serial written as the region types dates (2026-01-13 → "1/13/2026" in the US). */
function typedDate(serial: number, r: Regional): string {
  const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000)
  const [y, m, day] = [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]
  const parts = r.dateOrder === "dmy" ? [day, m, y] : r.dateOrder === "mdy" ? [m, day, y] : [y, m, day]
  return parts.join(r.dateSeparator)
}

export const REPLACE_DATA = "There's already data here. Do you want to replace it?"

/** A picture being reviewed: where it's shown from, and what was read in it. */
export type PictureReview = { url: string; read: PictureRead }

export class PictureData {
  /** The picture under review (the Data from Picture dialog shows it); null when none is. */
  review: PictureReview | null = null

  constructor(private readonly c: SheetController) {}

  /**
   * Reads a picture and opens the review. With text recognition off, asks to turn it on first (it then goes on). A
   * picture with no text is refused with a message a person can read.
   */
  async fromPicture(image: Blob): Promise<void> {
    const switcher = recognitionSwitch()
    if (!reader && switcher) {
      this.c.shell.confirm(
        "Read the picture's text with text recognition on this " +
          `${switcher.place}? Nothing is sent anywhere. You can turn it off again in File › Options › AI.`,
        () => {
          switcher.on()
          void this.fromPicture(image).catch((e: unknown) =>
            this.c.shell.alert(`Cells couldn't read the picture. ${message(e)}`),
          )
        },
      )
      return
    }
    if (!reader) throw new Error("Text recognition isn't available here.")
    const say = (text: string | null) => {
      this.c.getData.busy = text
      this.c.changed()
    }
    let read: PictureRead
    say("Recognizing text…")
    try {
      read = pictureTables(await reader(image))
    } finally {
      say(null)
    }
    if (!read.table.rows.length) throw new Error("Cells found no text in the picture.")
    this.close()
    this.review = { url: URL.createObjectURL(image), read }
    this.c.shell.openDialog?.("dataFromPicture")
  }

  /** Ends the review (Cancel, or after Insert Data). */
  close() {
    if (this.review) URL.revokeObjectURL(this.review.url)
    this.review = null
  }

  /**
   * Puts the rows into the sheet from the active cell, each as if typed there (numbers and dates read in the person's
   * region; a figure with a currency sign or spaced groups as its number; text that would be a formula stays text).
   * Asks first when that replaces data. Selects what went in; false when nothing did.
   */
  insert(rows: string[][]): boolean {
    if (this.c.edit) this.c.commitEdit()
    const width = Math.max(0, ...rows.map((r) => r.length))
    if (!rows.length || !width) return false
    const { row: r0, col: c0 } = this.c.selection.active
    const sheet = this.c.sheet
    const styles = this.c.wb.styles
    const cells: CellWrite[] = []
    let replaces = false
    const region = regional()
    const orders = Array.from({ length: width }, (_, k) =>
      dateOrderOf(
        rows.map((row) => row[k] ?? ""),
        region,
      ),
    )
    rows.forEach((r, i) => {
      for (let k = 0; k < width; k++) {
        const row = r0 + i
        const col = c0 + k
        const current = sheet.cells.get(row, col)
        if (current?.value !== undefined && current.value !== null && current.value !== "") replaces = true
        const text = (r[k] ?? "").trim()
        let data
        if (!text) data = current?.style ? { style: current.style } : undefined
        else if (/^[=+@]/.test(text) && Number.isNaN(Number(text))) data = { value: text, style: current?.style }
        else if (orders[k] && NUMERIC_DATE.test(text) && dateFrom(text, orders[k] as Regional) !== null)
          data = contentForInput(styles, current, typedDate(dateFrom(text, orders[k] as Regional) as number, region))
        else {
          data = contentForInput(styles, current, text)
          const n = typeof data?.value === "string" ? numberFrom(text) : null
          if (n !== null) data = { value: n, style: current?.style }
        }
        if (data && data.style === undefined) delete data.style
        cells.push({ row, col, data })
      }
    })
    const area = { r0, c0, r1: r0 + rows.length - 1, c1: c0 + width - 1 }
    const run = () => {
      const sheetIndex = this.c.sheetIndex
      const entry = this.c.transact([{ kind: "setCells", sheet: sheetIndex, cells }], {
        label: "Data from Picture",
        before: this.c.selection,
        after: selectRange(area, { row: r0, col: c0 }),
      })
      this.c.setSelection(selectRange(area, { row: r0, col: c0 }), null)
      // Columns too narrow for what went in (a date shown as ####) are widened; none is made narrower.
      const fit = this.c.columnFit
      const widths = []
      for (let col = area.c0; fit && entry && col <= area.c1; col++) {
        const size = fit(col, { r0: area.r0, r1: area.r1 }, sheetIndex)
        if (size !== null && size > this.c.sheet.cols.size(col)) widths.push({ index: col, size, hidden: false })
      }
      if (entry && widths.length)
        this.c.wb.amend(entry, [{ kind: "axis", sheet: sheetIndex, axis: "cols", entries: widths }])
    }
    if (replaces) this.c.shell.confirm(REPLACE_DATA, run)
    else run()
    return true
  }
}
