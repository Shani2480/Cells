import { DEFAULT_PROTECTION, PROTECTED_MESSAGE, parseRange, selectRange, type CellNote } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { CONVERT_MESSAGE } from "./annotations"
import { SHEET_MESSAGES } from "./sheets"

function setup() {
  const c = new SheetController()
  c.annotations.author = "Alex Demo"
  const messages: string[] = []
  const confirms: string[] = []
  c.shell = { ...c.shell, alert: (m) => messages.push(m), confirm: (m, ok) => (confirms.push(m), ok()) }
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  const at = (ref: string) => {
    const r = parseRange(ref)!
    return c.sheet.annotations.get(r.r0, r.c0)
  }
  return { c, a: c.annotations, select, at, messages, confirms }
}

describe("notes", () => {
  it("Shift+F2 adds a note and opens it; typing then leaving keeps the text, one undo step each", () => {
    const { c, a, select, at } = setup()
    select("B2")
    c.keyReady({ key: "F2", primary: false, shift: true, alt: false })
    expect(a.editing).toEqual({ row: 1, col: 1 })
    expect(at("B2")).toMatchObject({ kind: "note", author: "Alex Demo", text: "" })
    a.finishNote("Check the total")
    expect(a.editing).toBeNull()
    expect((at("B2") as CellNote).text).toBe("Check the total")
    c.wb.undo()
    expect((at("B2") as CellNote).text).toBe("")
    c.wb.undo()
    expect(at("B2")).toBeUndefined()
  })

  it("an unchanged note records nothing", () => {
    const { c, a, select } = setup()
    select("A1")
    a.newNote()
    a.finishNote("x")
    const depth = c.wb.undoLabel
    a.editNote()
    a.finishNote("x")
    expect(c.wb.undoLabel).toBe(depth)
  })

  it("shows and hides a note, and Delete Note removes the notes in the selection", () => {
    const { a, select, at } = setup()
    for (const ref of ["A1", "A3"]) {
      select(ref)
      a.newNote()
      a.finishNote(ref)
    }
    select("A1")
    a.toggleNote()
    expect((at("A1") as CellNote).shown).toBe(true)
    select("A1:A3")
    a.deleteNotes()
    expect(at("A1")).toBeUndefined()
    expect(at("A3")).toBeUndefined()
  })

  it("Next and Previous Note go round the sheet in reading order", () => {
    const { c, a, select } = setup()
    for (const ref of ["C1", "A5", "B2"]) {
      select(ref)
      a.newNote()
      a.finishNote(ref)
    }
    select("A1")
    a.step("note", 1)
    expect(c.selection.active).toEqual({ row: 0, col: 2 })
    a.finishNote("C1")
    a.step("note", 1)
    expect(c.selection.active).toEqual({ row: 1, col: 1 })
    a.finishNote("B2")
    a.step("note", -1)
    expect(c.selection.active).toEqual({ row: 0, col: 2 })
    a.finishNote("C1")
    a.step("note", -1)
    expect(c.selection.active).toEqual({ row: 4, col: 0 })
  })

  it("converts notes to comments after asking", () => {
    const { a, select, at, confirms } = setup()
    select("A1")
    a.newNote()
    a.finishNote("Old note")
    a.convertToComments()
    expect(confirms).toEqual([CONVERT_MESSAGE])
    expect(at("A1")).toMatchObject({ kind: "comment", posts: [{ author: "Alex Demo", text: "Old note" }] })
  })

  it("Clear All removes notes too", () => {
    const { c, a, select, at } = setup()
    select("A1")
    a.newNote()
    a.finishNote("x")
    c.format.clearAll()
    expect(at("A1")).toBeUndefined()
  })
})

describe("comments", () => {
  it("posts a thread, replies, edits, resolves and deletes", () => {
    const { a, select, at } = setup()
    select("D4")
    a.newComment()
    expect(a.open).toEqual({ row: 3, col: 3, draft: true })
    a.post("   ")
    expect(at("D4")).toBeUndefined()
    a.post("Is this right?")
    a.post("Yes")
    expect(at("D4")).toMatchObject({ kind: "comment", resolved: false })
    const thread = () => at("D4") as Extract<ReturnType<typeof at>, { kind: "comment" }>
    expect(thread().posts.map((p) => p.text)).toEqual(["Is this right?", "Yes"])
    a.editPost(3, 3, 1, "Yes, checked")
    expect(thread().posts[1]).toMatchObject({ text: "Yes, checked", edited: true })
    a.resolve(3, 3, true)
    expect(thread().resolved).toBe(true)
    a.deletePost(3, 3, 1)
    expect(thread().posts).toHaveLength(1)
    a.deletePost(3, 3, 0)
    expect(at("D4")).toBeUndefined()
    expect(a.open).toBeNull()
  })

  it("New Comment on a cell with a note opens the note instead", () => {
    const { a, select } = setup()
    select("A1")
    a.newNote()
    a.finishNote("n")
    a.newComment()
    expect(a.editing).toEqual({ row: 0, col: 0 })
    expect(a.open).toBeNull()
  })

  it("notes move with inserted rows", () => {
    const { c, a, select, at } = setup()
    select("A3")
    a.newNote()
    a.finishNote("n")
    select("A1")
    c.insertLines("rows")
    expect(at("A4")).toMatchObject({ text: "n" })
  })
})

describe("protect sheet", () => {
  const options = (() => {
    const { password: _p, ...rest } = DEFAULT_PROTECTION
    return rest
  })()

  it("refuses typing in a locked cell with Excel's message, and allows unlocked cells", () => {
    const { c, select, messages } = setup()
    select("B1")
    c.format.apply({ locked: false }, "Lock Cell")
    c.sheets.protect(options, "")
    select("A1")
    c.beginEdit("enter", "5")
    expect(c.edit).toBeNull()
    expect(messages).toEqual([PROTECTED_MESSAGE])
    select("B1")
    c.beginEdit("enter", "5")
    c.commitEdit()
    expect(c.sheet.cells.value(0, 1)).toBe(5)
  })

  it("refuses Delete, formatting and notes on locked cells", () => {
    const { c, a, select, messages } = setup()
    select("A1")
    c.beginEdit("enter", "1")
    c.commitEdit()
    c.sheets.protect(options, "")
    select("A1")
    c.clearContents()
    c.format.apply({ bold: true }, "Bold")
    a.newNote()
    expect(c.sheet.cells.value(0, 0)).toBe(1)
    expect(c.sheet.annotations.size).toBe(0)
    expect(messages).toHaveLength(3)
  })

  it("asks for the password to unprotect; a wrong one keeps the protection", () => {
    const { c, messages } = setup()
    c.sheets.protect(options, "s3cret")
    expect(c.sheets.needsPassword()).toBe(true)
    expect(c.sheets.unprotect("wrong")).toBe(false)
    expect(messages).toEqual([SHEET_MESSAGES.wrongPassword])
    expect(c.sheets.isProtected()).toBe(true)
    expect(c.sheets.unprotect("s3cret")).toBe(true)
    expect(c.sheets.isProtected()).toBe(false)
  })
})
