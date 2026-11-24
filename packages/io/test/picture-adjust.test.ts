import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { Workbook, type SheetObject, type SheetPicture } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(join(__dirname, "fixtures/excel/picture-adjust.xlsx")))
const pic = (list: SheetObject[], name: string) => list.find((o) => o.name === name) as SheetPicture
const parts = async (bytes: Uint8Array) => {
  const entries = await readZip(bytes)
  const text = (re: RegExp) => new TextDecoder().decode(entries.find((x) => re.test(x.name))?.data)
  return {
    names: entries.map((e) => e.name),
    drawing: text(/drawings\/drawing\d+\.xml$/),
    rels: text(/drawings\/_rels/),
  }
}
const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC"

describe("Picture Format › Corrections, Color and Transparency in Excel files", () => {
  it("reads Excel's corrections (with their HD Photo original) and its Office 2007 colour modes", async () => {
    const objects = (await openXlsx(fixture(), "x.xlsx")).workbook.sheets[0].objects
    const corrected = pic(objects, "Corrected")
    expect(corrected.adjust).toEqual({
      sharpen: 0.25,
      temperature: 8800,
      saturation: 1.33,
      bright: 0.2,
      contrast: -0.2,
    })
    expect(corrected.original?.startsWith("data:image/vnd.ms-photo;base64,")).toBe(true)
    expect(pic(objects, "Grayscale").adjust).toEqual({ recolor: { kind: "grayscale" } })
    expect(pic(objects, "BlackWhite").adjust).toEqual({ recolor: { kind: "blackWhite", threshold: 0.5 } })
    expect(pic(objects, "Washout").adjust).toEqual({ lum: { bright: 0.7, contrast: -0.7 } })
    expect(pic(objects, "TransparentColor").adjust).toEqual({ transparentColor: "787878" })
    expect(pic(objects, "Pencil").adjust).toEqual({ artistic: { effect: "PencilSketch" } })
  })

  it("writes Excel's pictures back with their originals, also when one is changed", async () => {
    const book = await openXlsx(fixture(), "x.xlsx")
    const corrected = pic(book.workbook.sheets[0].objects, "Corrected")
    corrected.adjust = { ...corrected.adjust, saturation: 2 }
    corrected.rotation = 10
    const saved = await saveXlsx(book.workbook, { source: book.source })
    const { drawing, rels, names } = await parts(saved)
    expect(drawing).toContain(`<a14:saturation sat="200000"/>`)
    expect(names.filter((n) => n.endsWith(".wdp"))).toHaveLength(1)
    expect(rels).toContain("relationships/hdphoto")
    const again = (await openXlsx(saved, "x.xlsx")).workbook.sheets[0].objects
    expect(pic(again, "Corrected").original).toBe(corrected.original)
    expect(pic(again, "Pencil").adjust).toEqual(pic(book.workbook.sheets[0].objects, "Pencil").adjust)
  })

  it("writes Cells' own: corrections with a PNG original, transparency, recolor and a transparent colour", async () => {
    const wb = new Workbook()
    const anchor = { from: { row: 1, col: 1, x: 0, y: 0 }, to: { row: 4, col: 3, x: 0, y: 0 } }
    const made: SheetPicture[] = [
      {
        kind: "picture",
        id: 1,
        name: "Picture 1",
        anchor,
        src: PNG,
        original: PNG,
        adjust: { bright: 0.4, contrast: -0.2, sharpen: -0.5, saturation: 0.66, temperature: 4700 },
      },
      {
        kind: "picture",
        id: 2,
        name: "Picture 2",
        anchor,
        src: PNG,
        adjust: {
          transparency: 0.5,
          recolor: { kind: "duotone", colors: [{ preset: "black" }, { theme: "accent2", tint: 0.45, satMod: 4 }] },
        },
      },
      {
        kind: "picture",
        id: 3,
        name: "Picture 3",
        anchor,
        src: PNG,
        adjust: { transparentColor: "FF00AA", lum: { bright: 0.7, contrast: -0.7 } },
      },
    ]
    wb.sheets[0].objects.push(...made)
    const saved = await saveXlsx(wb)
    const { drawing } = await parts(saved)
    expect(drawing).toContain(`<a:alphaModFix amt="50000"/>`)
    expect(drawing).toContain(
      `<a:duotone><a:prstClr val="black"/><a:schemeClr val="accent2"><a:tint val="45000"/><a:satMod val="400000"/></a:schemeClr></a:duotone>`,
    )
    const again = (await openXlsx(saved, "x.xlsx")).workbook.sheets[0].objects as SheetPicture[]
    expect(again.map((p) => p.adjust)).toEqual(made.map((p) => p.adjust))
    expect(again[0].original).toBe(PNG)
    expect(again[1].original).toBeUndefined()
  })

  it("reads Excel's artistic effects, their settings and their place among the corrections, and writes them back", async () => {
    const bytes = new Uint8Array(readFileSync(join(__dirname, "fixtures/excel/artistic-effects.xlsx")))
    const book = await openXlsx(bytes, "x.xlsx")
    const objects = book.workbook.sheets[0].objects
    expect(pic(objects, "Blur").adjust).toEqual({ artistic: { effect: "Blur", amount: 11 } })
    expect(pic(objects, "Cutout").adjust).toEqual({ artistic: { effect: "Cutout", trans: 0.4, amount: 3 } })
    expect(pic(objects, "Mosaic").adjust).toEqual({ artistic: { effect: "MosiaicBubbles" } })
    expect(pic(objects, "BrightThenBlur").adjust).toEqual({ bright: 0.2, artistic: { effect: "Blur", last: true } })
    // Excel's UI has one at a time; a second is kept as it was.
    const two = pic(objects, "TwoArtistic").adjust!
    expect(two.artistic).toEqual({ effect: "Blur" })
    expect(two.effects).toEqual([expect.stringMatching(/^<a14:artisticPaintBrush/)])

    const { drawing } = await parts(await saveXlsx(book.workbook, { source: book.source }))
    expect(drawing).toContain(`<a14:imgEffect><a14:artisticBlur radius="11"/></a14:imgEffect>`)
    expect(drawing).toContain(`<a14:artisticCutout trans="40000" numberOfShades="3"/>`)
    expect(drawing).toContain(`<a14:artisticMosiaicBubbles/>`)
    expect(drawing).toContain(
      `<a14:imgEffect><a14:brightnessContrast bright="20000"/></a14:imgEffect><a14:imgEffect><a14:artisticBlur/></a14:imgEffect>`,
    )
    // Office's defaults aren't written.
    const wb = new Workbook()
    wb.sheets[0].objects.push({
      kind: "picture",
      id: 1,
      name: "P",
      anchor: { from: { row: 0, col: 0, x: 0, y: 0 }, to: { row: 2, col: 2, x: 0, y: 0 } },
      src: PNG,
      original: PNG,
      adjust: { artistic: { effect: "CrisscrossEtching", trans: 0.75, amount: 50 }, saturation: 2 },
    })
    const mine = await parts(await saveXlsx(wb))
    expect(mine.drawing).toContain(
      `<a14:imgEffect><a14:artisticCrisscrossEtching pressure="50"/></a14:imgEffect><a14:imgEffect><a14:saturation`,
    )
  })
})
