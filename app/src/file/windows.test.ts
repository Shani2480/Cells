import { describe, expect, it } from "vitest"
import { newDocument } from "./document"

const setup = () => {
  const doc = newDocument()
  return { doc, w: doc.windows }
}

describe("View › New Window (windows as panels)", () => {
  it("New Window opens a second view of the same workbook, with its own place, and numbers the titles", () => {
    const { doc, w } = setup()
    const first = doc.controller
    first.scrollTo(0, 400)
    expect(doc.windowTitle).toBe("Book1")
    const second = w.newWindow()
    expect(doc.controller).toBe(second)
    expect(second.wb).toBe(first.wb)
    expect(second.scrollY).toBe(first.scrollY)
    expect(doc.windowTitle).toBe("Book1:2")
    expect(w.title(doc.title, first)).toBe("Book1:1")

    // Each window keeps its own selection; an edit in one shows in the other.
    second.typeInto({ row: 0, col: 0 }, "42", "Typing")
    expect(first.sheet.cells.value(0, 0)).toBe(42)
    second.scrollTo(0, 0)
    expect(first.scrollY).toBe(400)
    // Saving counts edits from every window.
    expect(doc.dirty).toBe(true)

    // Maximized (the default) only the active window shows; arranged, both do.
    expect(w.shown).toEqual([second])
    w.arrange("vertical")
    expect(w.shown).toEqual([first, second])

    w.close(second)
    expect(doc.controller).toBe(first)
    expect(doc.windowTitle).toBe("Book1")
    expect(w.arranged).toBe(false)
  })

  it("Hide and Unhide, Ctrl+F6, and closing renumbers what is left", () => {
    const { doc, w } = setup()
    const one = doc.controller
    const two = w.newWindow()
    const three = w.newWindow()
    w.arrange("tiled")
    w.hide()
    expect(w.hidden).toEqual([three])
    expect(w.active).toBe(two)
    expect(w.shown).toEqual([one, two])
    w.cycle(1)
    expect(w.active).toBe(one)
    w.unhide(three)
    expect(w.active).toBe(three)
    w.close(two)
    expect(w.title(doc.title, three)).toBe("Book1:2")
    // The last visible window can't be hidden.
    w.hide()
    w.hide()
    expect(w.visible).toHaveLength(1)
  })

  it("View Side by Side shows two windows one above the other, scrolling together until turned off", () => {
    const { doc, w } = setup()
    const one = doc.controller
    const two = w.newWindow()
    w.startSideBySide(one)
    expect(w.arrangement).toBe("horizontal")
    expect(w.shown).toEqual([one, two])

    // Two rows down in the active window: the other moves two rows too.
    two.scrollTo(0, two.sheet.rows.offset(2))
    expect(one.sheet.rows.indexAt(one.scrollY)).toBe(2)
    w.toggleSyncScroll()
    two.scrollTo(0, two.sheet.rows.offset(10))
    expect(one.sheet.rows.indexAt(one.scrollY)).toBe(2)

    w.endSideBySide()
    expect(w.sideBySide).toBeNull()
    expect(w.arranged).toBe(false)
  })
})
