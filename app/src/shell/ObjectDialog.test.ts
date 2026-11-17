import { isChart, OFFICE_THEME, type SheetPicture } from "@cells/core"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { embedSvg, fileKind, OBJECT_ICON_SIZE } from "@/drawing/embed"
import { objectSvg } from "@/drawing/render"
import { SheetController } from "@/sheet/controller"
import { objectPicture, progIdFor } from "./ObjectDialog"

/** An Image that loads at once (jsdom has no canvas: the SVG then stands in for the PNG). */
class FakeImage {
  onload: (() => void) | null = null
  set src(_: string) {
    queueMicrotask(() => this.onload?.())
  }
}

beforeEach(() => vi.stubGlobal("Image", FakeImage))
afterEach(() => vi.unstubAllGlobals())

describe("Insert › Object", () => {
  it("embeds a file as an icon with its name, one undo step, drawn in the colour of its kind", async () => {
    const c = new SheetController()
    const embed = { name: "Budget notes.docx", progId: progIdFor("Budget notes.docx"), data: new Uint8Array([1, 2]) }
    expect(embed.progId).toBe("Word.Document.12")
    expect(progIdFor("list.txt")).toBe("Package")
    // No canvas in tests: the SVG stands in for the PNG.
    const src = await objectPicture(embed, c)
    expect(src).toMatch(/^data:image\/(png|svg\+xml);base64,/)
    c.objects.insertObject(embed, src, OBJECT_ICON_SIZE)
    const o = c.sheet.objects.find((x): x is SheetPicture => !isChart(x) && x.kind === "picture")!
    expect(o.embed).toEqual(embed)
    expect(o.name).toMatch(/^Object \d+$/)
    expect(c.wb.undoLabel).toBe("Insert Object")
    const svg = objectSvg(o, OBJECT_ICON_SIZE.w, OBJECT_ICON_SIZE.h, OFFICE_THEME)
    expect(svg).toContain(">W</text>")
    expect(svg).toMatch(/>Budget<\/text>.*>notes\.docx<\/text>/)
    expect(embedSvg({ name: "Quarterly budget notes.docx", progId: "" }, 84, 76, OFFICE_THEME)).toMatch(
      />Quarterly<\/text>.*>budget…es\.docx<\/text>/,
    )
    c.wb.undo()
    expect(c.sheet.objects).toHaveLength(0)
  })

  it("knows Office, PDF and picture files, and shortens long names in the middle", () => {
    expect(fileKind("a.xlsx")?.mark).toBe("X")
    expect(fileKind("a.PDF")?.mark).toBe("PDF")
    expect(fileKind("photo.jpeg")?.mark).toBe("IMG")
    expect(fileKind("thing.bin")).toBeNull()
    expect(fileKind("Embedded", "PowerPoint.Show.12")?.mark).toBe("P")
    const svg = embedSvg(
      { name: "A very long file name for the quarterly report.txt", progId: "Package" },
      84,
      76,
      OFFICE_THEME,
    )
    expect(svg).toMatch(/>A very long<\/text>.*>file na…ort\.txt<\/text>/)
  })
})
