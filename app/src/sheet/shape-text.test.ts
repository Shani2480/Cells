import { describe, expect, it } from "vitest"
import { plainText, textFromPlain, type ShapeText } from "@cells/core"
import { SheetController } from "./controller"

/**
 * Home › Font and Alignment on text in shapes, as in Excel: with a shape selected they format all its text; while
 * it's being typed in, the selected text only. The cells keep their own format meanwhile.
 */
function setup() {
  const c = new SheetController()
  const shape = c.objects.insertShape("rect")
  c.objects.setText(shape.id, "Sales up\nNext year")
  c.objects.select(shape.id)
  const text = () => (c.objects.find(shape.id) as { text: ShapeText }).text
  return { c, id: shape.id, text }
}

describe("formatting text in shapes", () => {
  it("a selected shape: Bold, size, font, colour and alignment format all its text, one undo step each", () => {
    const { c, text } = setup()
    c.format.toggle("bold")
    expect(text().paragraphs.every((p) => p.runs.every((r) => r.bold))).toBe(true)
    expect(c.format.active.bold).toBe(true)
    c.format.fontSize(20)
    c.format.fontName("Georgia")
    c.format.fontColor("#c00000")
    expect(text().paragraphs[1].runs[0]).toMatchObject({ size: 20, font: "Georgia", color: { rgb: "C00000" } })
    expect(c.format.active).toMatchObject({ fontSize: 20, fontName: "Georgia", fontColor: "C00000" })
    // New shapes centre their text; the buttons choose one alignment.
    expect(c.format.active.hAlign).toBe("center")
    c.format.hAlign("right")
    c.format.vAlign("bottom")
    expect(text().paragraphs.map((p) => p.align)).toEqual(["r", "r"])
    expect(text().anchor).toBe("b")
    expect(c.format.active).toMatchObject({ hAlign: "right", vAlign: "bottom" })
    // The cell under it kept its own format.
    expect(c.wb.styles.get(c.sheet.styleAt(0, 0)).bold).toBeUndefined()
    c.undo()
    c.undo()
    expect(text().anchor).not.toBe("b")
    expect(text().paragraphs[0].align).toBe("ctr")
    c.objects.select(null)
    expect(c.format.active.bold).toBeUndefined()
  })

  it("while typing: the editor's selection only, kept with the text when it closes", () => {
    const { c, id, text } = setup()
    let draft = text()
    let range: [number, number] = [6, 8]
    c.objects.editing = id
    c.objects.textEditor = { text: () => draft, range: () => range, set: (t) => (draft = t) }
    expect(c.objects.formatsText()).toBe(true)
    c.format.toggle("italic")
    c.format.underline()
    expect(draft.paragraphs[0].runs.map((r) => [r.text, !!r.italic, !!r.underline])).toEqual([
      ["Sales ", false, false],
      ["up", true, true],
    ])
    expect(c.format.active.italic).toBe(true)
    range = [9, 9]
    expect(c.format.active.italic).toBeUndefined()
    // A caret alone formats nothing; alignment takes its paragraph.
    c.format.toggle("bold")
    expect(draft.paragraphs.some((p) => p.runs.some((r) => r.bold))).toBe(false)
    c.format.hAlign("right")
    expect(draft.paragraphs.map((p) => p.align)).toEqual(["ctr", "r"])
    c.objects.textEditor = null
    c.objects.setRichText(id, draft)
    expect(text().paragraphs[0].runs[1]).toMatchObject({ text: "up", italic: true })
    expect(plainText(text())).toBe("Sales up\nNext year")
    c.undo()
    expect(text().paragraphs[0].runs).toHaveLength(1)
  })

  it("an empty shape keeps the look for what's typed in later; lines take none", () => {
    const c = new SheetController()
    const shape = c.objects.insertShape("rect")
    c.objects.select(shape.id)
    c.format.fontSize(24)
    const t = (c.objects.find(shape.id) as { text: ShapeText }).text
    expect(t.paragraphs[0].runs).toEqual([expect.objectContaining({ text: "", size: 24 })])
    expect(textFromPlain("Hi", t).paragraphs[0].runs[0].size).toBe(24)
    const line = c.objects.insertShape("line")
    c.objects.select(line.id)
    expect(c.objects.formatsText()).toBe(false)
  })
})
