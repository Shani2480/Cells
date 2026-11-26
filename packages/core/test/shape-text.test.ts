import { describe, expect, it } from "vitest"
import { alignText, formatText, lookAt, replaceText, textLength, textPositions, type ShapeText } from "../src"

const text = (): ShapeText => ({
  paragraphs: [{ runs: [{ text: "Hello world", size: 14 }] }, { runs: [] }, { align: "ctr", runs: [{ text: "Bye" }] }],
})
const show = (t: ShapeText) =>
  t.paragraphs
    .map((p) => p.runs.map((r) => `${r.bold ? "*" : ""}${r.italic ? "/" : ""}${r.text}`).join("|"))
    .join(" ¶ ")

describe("formatting part of a shape's text", () => {
  it("splits runs at the range and merges equal neighbours", () => {
    const t = formatText(text(), 6, 11, { bold: true })
    expect(show(t)).toBe("Hello |*world ¶  ¶ Bye")
    expect(t.paragraphs[0].runs[1].size).toBe(14)
    expect(show(formatText(t, 0, 6, { bold: true }))).toBe("*Hello world ¶  ¶ Bye")
    expect(show(formatText(t, 6, 11, { bold: null }))).toBe("Hello world ¶  ¶ Bye")
  })

  it("spans paragraphs (the break counts as one), giving an empty paragraph the look", () => {
    expect(textLength(text())).toBe(16)
    const t = formatText(text(), 3, 15, { italic: true })
    expect(show(t)).toBe("Hel|/lo world ¶ / ¶ /By|e")
    expect(t.paragraphs[1].runs).toEqual([{ text: "", italic: true }])
  })

  it("the ribbon's look: a selection's first character, a caret's one before", () => {
    const t = formatText(text(), 6, 11, { bold: true })
    expect(lookAt(t, 6, 8).bold).toBe(true)
    expect(lookAt(t, 6).bold).toBeUndefined()
    expect(lookAt(t, 7).bold).toBe(true)
    expect(lookAt(t, 0).size).toBe(14)
  })

  it("aligns the paragraphs a range touches", () => {
    const t = alignText(text(), 2, 2, "r")
    expect(t.paragraphs.map((p) => p.align)).toEqual(["r", undefined, "ctr"])
    expect(alignText(text(), 0, 16, null).paragraphs.map((p) => p.align)).toEqual([undefined, undefined, undefined])
  })
})

describe("replacing a word in shape text", () => {
  const text: ShapeText = {
    paragraphs: [
      { runs: [{ text: "Plan " }, { text: "teh", bold: true }, { text: " year" }] },
      { runs: [{ text: "wrnog" }] },
    ],
  }

  it("keeps the look of the run the word is in, and the runs around it", () => {
    expect(textPositions(text)).toBe("Plan teh year\nwrnog")
    const out = replaceText(text, 5, 8, "the")
    expect(out.paragraphs[0].runs).toEqual([{ text: "Plan " }, { text: "the", bold: true }, { text: " year" }])
    expect(replaceText(text, 14, 19, "wrong").paragraphs[1].runs).toEqual([{ text: "wrong" }])
    // A word spanning two runs takes the first run's look.
    const split: ShapeText = { paragraphs: [{ runs: [{ text: "re", italic: true }, { text: "cieve" }] }] }
    expect(replaceText(split, 0, 7, "receive").paragraphs[0].runs).toEqual([{ text: "receive", italic: true }])
  })
})
