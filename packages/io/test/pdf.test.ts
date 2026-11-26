// The PDF writer: one page per picture, with a cross-reference table that points at every object.
import { describe, expect, it } from "vitest"
import { pdfFromImages } from "../src"
import { jpegColorSpace } from "../src/pdf"
import { glyphlessFont } from "../src/pdf-font"

// The smallest JPEG: a 1×1 white picture.
const JPEG = Uint8Array.from(
  atob(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AKp//2Q==",
  ),
  (c) => c.charCodeAt(0),
)

describe("PDF of pictures", () => {
  it("writes a page per picture at its size, every object where the table says", () => {
    const bytes = pdfFromImages(
      [
        { width: 595.28, height: 841.89, jpeg: JPEG, pixelWidth: 1, pixelHeight: 1 },
        { width: 841.89, height: 595.28, jpeg: JPEG, pixelWidth: 1, pixelHeight: 1 },
      ],
      { title: "Report – Q1", date: new Date(Date.UTC(2026, 10, 15, 9, 30)) },
    )
    const text = new TextDecoder("latin1").decode(bytes)
    expect(text.startsWith("%PDF-1.4")).toBe(true)
    expect(text).toContain("/Count 2")
    expect(text).toContain("/MediaBox [0 0 595.28 841.89]")
    expect(text).toContain("/MediaBox [0 0 841.89 595.28]")
    expect(text).toContain("/CreationDate (D:20261115093000Z)")
    // Non-ASCII titles are UTF-16.
    expect(text).toContain("/Title <FEFF005200650070006F0072007400202013002000510031>")
    const xref = Number(/startxref\n(\d+)/.exec(text)![1])
    expect(text.slice(xref, xref + 4)).toBe("xref")
    const entries = text.slice(xref).split("\n").slice(3, 12)
    entries.forEach((line, i) => {
      const at = Number(line.slice(0, 10))
      expect(text.slice(at, at + `${i + 1} 0 obj`.length)).toBe(`${i + 1} 0 obj`)
    })
  })

  it("declares a picture's colours as its JPEG has them (grey scans stay grey)", () => {
    // The 1×1 picture above is grey: its frame header has one channel.
    expect(jpegColorSpace(JPEG)).toBe("/DeviceGray")
    const rgb = JPEG.slice()
    rgb[rgb.indexOf(0xc0, 2) + 8] = 3
    expect(jpegColorSpace(rgb)).toBe("/DeviceRGB")
    const text = new TextDecoder("latin1").decode(
      pdfFromImages([{ width: 10, height: 10, jpeg: JPEG, pixelWidth: 1, pixelHeight: 1 }]),
    )
    expect(text).toContain("/ColorSpace /DeviceGray")
  })

  it("lays a page's text over its picture invisibly, in a shapeless font that maps back to the characters", () => {
    const page = (text: { text: string; x: number; y: number; size: number; width: number; angle?: number }[]) => ({
      width: 612,
      height: 792,
      jpeg: JPEG,
      pixelWidth: 1,
      pixelHeight: 1,
      text,
    })
    const bytes = pdfFromImages([
      page([
        { text: "Sales", x: 72, y: 100, size: 11, width: 25 },
        { text: "Up 😀", x: 300, y: 400, size: 10, width: 30, angle: -90 },
      ]),
      page([{ text: "Sales", x: 72, y: 72, size: 11, width: 25 }]),
    ])
    const text = new TextDecoder("latin1").decode(bytes)
    // Render mode 3 (invisible), the text stretched to its width (5 glyphs half an em wide at 11 pt → 25 pt).
    expect(text).toContain("BT 3 Tr /F1 1 Tf\n10 0 0 11 72 692 Tm <00010002000300040005> Tj")
    // Turned up the page: the baseline runs upwards.
    expect(text).toContain("0 15 -10 0 300 392 Tm <0006000700080009> Tj")
    expect(text.match(/\/Font << \/F1 \d+ 0 R >>/g)).toHaveLength(2)
    expect(text).toContain("/Subtype /Type0 /BaseFont /GlyphLessFont /Encoding /Identity-H")
    // Each character once, the emoji as its UTF-16 pair.
    expect(text).toContain("9 beginbfchar\n<0001> <0053>\n<0002> <0061>")
    expect(text).toContain("<0009> <D83DDE00>")
    // The cross-reference table still finds every object.
    const xref = Number(/startxref\n(\d+)/.exec(text)![1])
    const count = Number(/\/Size (\d+)/.exec(text)![1])
    text
      .slice(xref)
      .split("\n")
      .slice(3, 2 + count)
      .forEach((line, i) => {
        const at = Number(line.slice(0, 10))
        expect(text.slice(at, at + `${i + 1} 0 obj`.length)).toBe(`${i + 1} 0 obj`)
      })
  })

  it("leaves the font out of a PDF without text", () => {
    const text = new TextDecoder("latin1").decode(
      pdfFromImages([{ width: 10, height: 10, jpeg: JPEG, pixelWidth: 1, pixelHeight: 1, text: [] }]),
    )
    expect(text).not.toContain("/Font")
    expect(text).not.toContain(" Tr ")
  })

  it("makes a valid TrueType font: its tables' checksums and the whole font's add up", () => {
    const font = glyphlessFont()
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength)
    const sum = (from: number, length: number) => {
      let s = 0
      for (let i = from; i < from + length; i += 4) s = (s + view.getUint32(i)) >>> 0
      return s
    }
    const tables = view.getUint16(4)
    for (let i = 0; i < tables; i++) {
      const at = 12 + i * 16
      const tag = String.fromCharCode(...font.slice(at, at + 4))
      const offset = view.getUint32(at + 8)
      const length = view.getUint32(at + 12)
      if (tag !== "head") expect(sum(offset, (length + 3) & ~3), tag).toBe(view.getUint32(at + 4))
    }
    expect(sum(0, font.length)).toBe(0xb1b0afba)
  })
})
