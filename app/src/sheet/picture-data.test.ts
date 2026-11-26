import { EN_US, selectCell, setRegional, type Regional } from "@cells/core"
import { afterEach, describe, expect, it, vi } from "vitest"
import type { PictureReader } from "@/pdf/picture"
import { SheetController } from "./controller"
import { setRecognitionSwitch } from "./get-data"
import { dateOrderOf, REPLACE_DATA, setPictureReader } from "./picture-data"

const GB: Regional = { decimal: ".", group: ",", dateSeparator: "/", dateOrder: "dmy", shortDate: "dd/mm/yyyy" }
const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000

/** A recognizer's answer for a little table: a word per cell, 40 pixels a column, one marked unsure. */
const reader: PictureReader = async () =>
  [
    ["Item", "Qty", "Price"],
    ["A-210", "500", "€0.12"],
    ["C-300", "2", "€18.50"],
  ].map((cells, r) => {
    const y0 = 100 + r * 40
    const words = cells.map((text, c) => ({
      text,
      confidence: r === 2 && c === 1 ? 40 : 95,
      bbox: { x0: 100 + c * 200, y0, x1: 100 + c * 200 + text.length * 14, y1: y0 + 24 },
    }))
    return { bbox: { x0: 100, y0, x1: words[words.length - 1].bbox.x1, y1: y0 + 24 }, words }
  })

afterEach(() => {
  setPictureReader(null)
  setRecognitionSwitch(null)
  setRegional(EN_US)
})

describe("Data › From Picture", () => {
  it("reads a picture for review, marking what recognition wasn't sure of", async () => {
    const c = new SheetController()
    const opened: string[] = []
    c.shell = { ...c.shell, openDialog: (name) => void opened.push(name) }
    setPictureReader(reader)
    globalThis.URL.createObjectURL ??= () => "blob:picture"
    globalThis.URL.revokeObjectURL ??= () => undefined
    await c.pictureData.fromPicture(new Blob(["x"], { type: "image/png" }))
    expect(opened).toEqual(["dataFromPicture"])
    expect(c.pictureData.review?.read.table.rows).toEqual([
      ["Item", "Qty", "Price"],
      ["A-210", "500", "€0.12"],
      ["C-300", "2", "€18.50"],
    ])
    expect(c.pictureData.review?.read.table.unsure).toEqual([[2, 1]])
    expect(c.getData.busy).toBeNull()
  })

  it("asks to turn text recognition on first, then goes on", async () => {
    const c = new SheetController()
    const asked: string[] = []
    c.shell = { ...c.shell, confirm: (m, ok) => (asked.push(m), ok()), openDialog: () => undefined }
    setRecognitionSwitch(() => setPictureReader(reader), "device")
    await c.pictureData.fromPicture(new Blob(["x"], { type: "image/png" }))
    expect(asked[0]).toMatch(/text recognition on this device\?/)
    await vi.waitFor(() => expect(c.pictureData.review).not.toBeNull())
  })

  it("puts the cells in at the active cell as typed in the region, in one step to undo", () => {
    setRegional(GB)
    const c = new SheetController()
    c.setSelection(selectCell({ row: 1, col: 1 }), null)
    expect(
      c.pictureData.insert([
        ["Date", "Amount", "Note"],
        ["03/01/2026", "€ 450.00", "=SUM(A1)"],
        ["04/01/2026", "1,200.50", ""],
      ]),
    ).toBe(true)
    const v = (r: number, col: number) => c.sheet.cells.get(r, col)?.value
    expect(v(1, 1)).toBe("Date")
    expect(v(2, 1)).toBe(serial(2026, 1, 3))
    expect(v(2, 2)).toBe(450)
    expect(v(2, 3)).toBe("=SUM(A1)")
    expect(c.sheet.cells.get(2, 3)?.formula).toBeUndefined()
    expect(v(3, 2)).toBe(1200.5)
    expect(c.selection.ranges[0]).toEqual({ r0: 1, c0: 1, r1: 3, c1: 3 })
    c.undo()
    expect(v(1, 1)).toBeUndefined()
    expect(v(2, 2)).toBeUndefined()
  })

  it("reads a column of dates in the order they're all written in, when that isn't the region's", () => {
    expect(dateOrderOf(["Date", "01/01/2026", "12/01/2026", "13/01/2026"], EN_US)?.dateOrder).toBe("dmy")
    expect(dateOrderOf(["Date", "01/02/2026", "03/04/2026"], EN_US)).toBeNull()
    expect(dateOrderOf(["Note", "Coffee", "13/01/2026"], EN_US)).toBeNull()
    const c = new SheetController()
    c.pictureData.insert([
      ["Date", "Amount"],
      ["12/01/2026", "-5.00"],
      ["13/01/2026", "-6.00"],
    ])
    expect(c.sheet.cells.get(1, 0)?.value).toBe(serial(2026, 1, 12))
    expect(c.sheet.cells.get(2, 0)?.value).toBe(serial(2026, 1, 13))
    expect(c.wb.styles.get(c.sheet.styleAt(2, 0)).numberFormat).toMatch(/[dy]/)
  })

  it("asks before replacing data", () => {
    const c = new SheetController()
    c.typeInto({ row: 0, col: 0 }, "Kept", "Typing")
    const asked: string[] = []
    c.shell = { ...c.shell, confirm: (m) => void asked.push(m) }
    c.pictureData.insert([["New"]])
    expect(asked).toEqual([REPLACE_DATA])
    expect(c.sheet.cells.get(0, 0)?.value).toBe("Kept")
  })
})
