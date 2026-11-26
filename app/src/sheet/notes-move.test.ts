import { parseRange, selectRange, type CellNote } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"

/**
 * Notes and comments stay with their cells, as measured in Excel 365: Insert / Delete Cells move them, cut and paste
 * moves them (replacing the destination's), copy and paste copies them (also replacing), sort moves them with their
 * rows; Fill and pasting text from elsewhere leave them.
 */
function setup() {
  const c = new SheetController()
  const note = (text: string): CellNote => ({ kind: "note", author: "Alex Demo", text, shown: false })
  const at = (ref: string) => {
    const r = parseRange(ref)!
    return (c.sheet.annotations.get(r.r0, r.c0) as CellNote | undefined)?.text
  }
  const put = (ref: string, text: string, value: string | number = text) => {
    const r = parseRange(ref)!
    c.sheet.cells.set(r.r0, r.c0, { value })
    c.sheet.annotations.set(r.r0, r.c0, note(text))
  }
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  // A1:A3 = 3, 1, 2; B1:B3 = c, a, b with notes n3, n1, n2 (the Excel check's sheet).
  c.sheet.cells.set(0, 0, { value: 3 })
  c.sheet.cells.set(1, 0, { value: 1 })
  c.sheet.cells.set(2, 0, { value: 2 })
  put("B1", "n3", "c")
  put("B2", "n1", "a")
  put("B3", "n2", "b")
  const notes = () =>
    c.sheet.annotations
      .all()
      .sort((x, y) => x.row - y.row || x.col - y.col)
      .map((a) => `${String.fromCharCode(65 + a.col)}${a.row + 1}=${(a.value as CellNote).text}`)
      .join(", ")
  return { c, at, put, select, notes }
}

const START = "B1=n3, B2=n1, B3=n2"

describe("notes and comments move with their cells", () => {
  it("Insert Cells shifting down or right moves them; undo puts them back", () => {
    const { c, select, notes } = setup()
    select("B1")
    c.shiftCells("insert", "down")
    expect(notes()).toBe("B2=n3, B3=n1, B4=n2")
    c.undo()
    expect(notes()).toBe(START)
    select("A2")
    c.shiftCells("insert", "right")
    expect(notes()).toBe("B1=n3, C2=n1, B3=n2")
  })

  it("Delete Cells shifting up removes the deleted cell's note and moves the rest; undo brings it back", () => {
    const { c, select, notes } = setup()
    select("B1")
    c.shiftCells("delete", "up")
    expect(notes()).toBe("B1=n1, B2=n2")
    c.undo()
    expect(notes()).toBe(START)
    select("A2")
    c.shiftCells("delete", "left")
    expect(notes()).toBe("B1=n3, A2=n1, B3=n2")
  })

  it("cut and paste moves them, replacing the destination's; undo restores both", () => {
    const { c, put, select, notes } = setup()
    put("D5", "old", "x")
    select("B2")
    const cut = c.copy(true)!
    select("D5")
    c.paste(cut.text, cut.html)
    expect(notes()).toBe("B1=n3, B3=n2, D5=n1")
    c.undo()
    expect(notes()).toBe("B1=n3, B2=n1, B3=n2, D5=old")
    // A cell without a note, cut onto one with a note, takes it away.
    select("A2")
    const plain = c.copy(true)!
    select("D5")
    c.paste(plain.text, plain.html)
    expect(notes()).toBe(START)
  })

  it("a cut pasted on another sheet takes them there", () => {
    const { c, select, notes } = setup()
    select("B1:B2")
    const cut = c.copy(true)!
    const first = c.sheet
    c.sheets.insert()
    expect(c.sheet).not.toBe(first)
    select("C3")
    c.paste(cut.text, cut.html)
    expect(notes()).toBe("C3=n3, C4=n1")
    expect(first.annotations.all().map((a) => (a.value as CellNote).text)).toEqual(["n2"])
  })

  it("copy and paste copies them (tiled with the cells) and replaces the destination's", () => {
    const { c, put, select, notes, at } = setup()
    select("B2:B3")
    const copy = c.copy(false)!
    select("D1:D4")
    c.paste(copy.text, copy.html)
    expect(notes()).toBe("B1=n3, D1=n1, B2=n1, D2=n2, B3=n2, D3=n1, D4=n2")
    put("F5", "old", "x")
    select("A2")
    const plain = c.copy(false)!
    select("F5")
    c.paste(plain.text, plain.html)
    expect(at("F5")).toBeUndefined()
    // The pasted note is a copy: changing one leaves the other.
    ;(c.sheet.annotations.get(0, 3) as CellNote).text = "changed"
    expect(at("B2")).toBe("n1")
  })

  it("text pasted from elsewhere, and Fill Down, leave them", () => {
    const { c, select, notes } = setup()
    select("B2")
    c.paste("from another app", null)
    expect(notes()).toBe(START)
    select("B2:B5")
    c.fillFromEdge("down")
    expect(c.sheet.cells.get(2, 1)?.value).toBe("from another app")
    expect(notes()).toBe(START)
  })

  it("Sort moves them with their rows", () => {
    const { c, notes } = setup()
    c.data.sort(parseRange("A1:B3")!, { header: false, keys: [{ index: 0, descending: false }] })
    expect(notes()).toBe("B1=n1, B2=n2, B3=n3")
  })
})
