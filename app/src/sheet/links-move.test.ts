import { parseRange, selectRange, type Hyperlink } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"

/**
 * Links stay with their cells, as measured in Excel 365 (the same sheet there and here): B2 → one, B4:C4 → wide,
 * E2 → target, H2 plain text. A copy takes the links on its cells and replaces the destination's (a plain cell
 * removes one); a cut moves only links wholly inside it; Insert / Delete Cells move a link wholly in the shifted band
 * and leave one reaching outside it.
 */
function setup() {
  const c = new SheetController()
  const link = (ref: string, name: string): Hyperlink => ({
    range: parseRange(ref)!,
    address: `https://example.com/${name}`,
  })
  const text = (ref: string, value: string) => {
    const r = parseRange(ref)!
    for (let row = r.r0; row <= r.r1; row++)
      for (let col = r.c0; col <= r.c1; col++) c.sheet.cells.set(row, col, { value })
  }
  text("B2", "one")
  text("B4:C4", "wide")
  text("E2", "target")
  text("H2", "plain")
  c.sheet.hyperlinks = [link("B2", "one"), link("B4:C4", "wide"), link("E2", "target")]
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  const ref = (r: Hyperlink["range"]) => {
    const a = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`
    return r.r0 === r.r1 && r.c0 === r.c1 ? a(r.r0, r.c0) : `${a(r.r0, r.c0)}:${a(r.r1, r.c1)}`
  }
  const links = (sheet = c.sheet) =>
    sheet.hyperlinks
      .map((l) => `${ref(l.range)}=${l.address!.replace("https://example.com/", "")}`)
      .sort()
      .join(" ")
  const copyTo = (from: string, to: string, cut = false) => {
    select(from)
    const clip = c.copy(cut)!
    select(to)
    c.paste(clip.text, clip.html)
  }
  return { c, select, links, copyTo }
}

const START = "B2=one B4:C4=wide E2=target"

describe("links move with their cells (as Excel 365)", () => {
  it("copy and paste takes the link along, trimmed to the cells copied, and tiles it", () => {
    const { c, links, copyTo } = setup()
    copyTo("B2", "D6")
    expect(links()).toBe("B2=one B4:C4=wide D6=one E2=target")
    c.undo()
    expect(links()).toBe(START)
    copyTo("B4", "B8")
    expect(links()).toBe(`B2=one B4:C4=wide B8=wide E2=target`)
    c.undo()
    copyTo("B2", "J2:K3")
    expect(links()).toBe(`${START} J2=one J3=one K2=one K3=one`.split(" ").sort().join(" "))
  })

  it("a pasted cell replaces the destination's link, and a plain one removes it", () => {
    const { links, copyTo, c } = setup()
    copyTo("B2", "E2")
    expect(links()).toBe("B2=one B4:C4=wide E2=one")
    c.undo()
    copyTo("H2", "E2")
    expect(links()).toBe("B2=one B4:C4=wide")
  })

  it("cut and paste moves a link wholly inside the cut, replacing the destination's; undo restores both", () => {
    const { c, links, copyTo } = setup()
    copyTo("B2", "F6", true)
    expect(links()).toBe("B4:C4=wide E2=target F6=one")
    c.undo()
    expect(links()).toBe(START)
    copyTo("B2", "E2", true)
    expect(links()).toBe("B4:C4=wide E2=one")
    c.undo()
    expect(links()).toBe(START)
    // Half of a wider link: it stays where it is.
    copyTo("B4", "B9", true)
    expect(links()).toBe(START)
  })

  it("a cut pasted on another sheet takes its links there", () => {
    const { c, select, links } = setup()
    select("B2")
    const cut = c.copy(true)!
    const first = c.sheet
    c.sheets.insert()
    select("A1")
    c.paste(cut.text, cut.html)
    expect(links()).toBe("A1=one")
    expect(links(first)).toBe("B4:C4=wide E2=target")
  })

  it("Insert Cells moves links wholly in the shifted band and leaves those reaching outside it", () => {
    const { c, select, links } = setup()
    select("B1")
    c.shiftCells("insert", "down")
    expect(links()).toBe("B3=one B4:C4=wide E2=target")
    c.undo()
    expect(links()).toBe(START)
    select("B1:C1")
    c.shiftCells("insert", "down")
    expect(links()).toBe("B3=one B5:C5=wide E2=target")
    c.undo()
    select("A2")
    c.shiftCells("insert", "right")
    expect(links()).toBe("B4:C4=wide C2=one F2=target")
  })

  it("Delete Cells removes a deleted cell's link, shrinks or moves the rest; undo brings them back", () => {
    const { c, select, links } = setup()
    select("B1")
    c.shiftCells("delete", "up")
    expect(links()).toBe("B1=one B4:C4=wide E2=target")
    c.undo()
    select("B2")
    c.shiftCells("delete", "up")
    expect(links()).toBe("B4:C4=wide E2=target")
    c.undo()
    expect(links()).toBe(START)
    select("A4")
    c.shiftCells("delete", "left")
    expect(links()).toBe("A4:B4=wide B2=one E2=target")
    c.undo()
    select("B4")
    c.shiftCells("delete", "left")
    expect(links()).toBe("B2=one B4=wide E2=target")
    c.undo()
    expect(links()).toBe(START)
  })
})
