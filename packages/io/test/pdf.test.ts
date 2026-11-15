// The PDF writer: one page per picture, with a cross-reference table that points at every object.
import { describe, expect, it } from "vitest"
import { pdfFromImages } from "../src"

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
})
