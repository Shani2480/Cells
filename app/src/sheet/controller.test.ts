import { describe, expect, it } from "vitest"
import { cellName, rangeName } from "@cells/core"
import { SheetController, type KeyInput } from "./controller"

const k = (key: string, mods: Partial<Omit<KeyInput, "key">> = {}): KeyInput => ({
  key,
  primary: false,
  shift: false,
  alt: false,
  ...mods,
})

function setup() {
  const c = new SheetController()
  c.viewWidth = 1000
  c.viewHeight = 600
  return c
}

/** Types text into the active cell the way the view does: the first character opens an entry. */
function type(c: SheetController, text: string) {
  c.beginEdit("enter", text)
}

function press(c: SheetController, key: string, mods: Partial<Omit<KeyInput, "key">> = {}) {
  const input = k(key, mods)
  return c.edit ? c.keyEditing(input) : c.keyReady(input)
}

const active = (c: SheetController) => cellName(c.selection.active)
const value = (c: SheetController, ref: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
  return c.sheet.cells.value(Number(m[2]) - 1, col)
}

describe("entering data", () => {
  it("types, commits with Enter and moves down", () => {
    const c = setup()
    type(c, "42")
    expect(c.mode).toBe("enter")
    press(c, "Enter")
    expect(value(c, "A1")).toBe(42)
    expect(active(c)).toBe("A2")
    expect(c.wb.undoLabel).toBe("Typing '42' in A1")
  })

  it("Tab moves right and Enter returns to the column where the Tab run started", () => {
    const c = setup()
    press(c, "ArrowRight")
    type(c, "a")
    press(c, "Tab")
    type(c, "b")
    press(c, "Tab")
    type(c, "c")
    press(c, "Enter")
    expect(active(c)).toBe("B2")
    expect([value(c, "B1"), value(c, "C1"), value(c, "D1")]).toEqual(["a", "b", "c"])
  })

  it("Esc cancels the entry and keeps the old value", () => {
    const c = setup()
    type(c, "1")
    press(c, "Enter")
    press(c, "ArrowUp")
    type(c, "2")
    press(c, "Escape")
    expect(c.edit).toBeNull()
    expect(value(c, "A1")).toBe(1)
  })

  it("Ctrl+Enter fills every selected cell and keeps the selection", () => {
    const c = setup()
    press(c, "ArrowDown", { shift: true })
    press(c, "ArrowRight", { shift: true })
    type(c, "x")
    press(c, "Enter", { primary: true })
    expect(rangeName(c.selection.ranges[0])).toBe("A1:B2")
    expect([value(c, "A1"), value(c, "B1"), value(c, "A2"), value(c, "B2")]).toEqual(["x", "x", "x", "x"])
  })

  it("Enter inside a selection walks it instead of leaving it", () => {
    const c = setup()
    press(c, "ArrowDown", { shift: true })
    press(c, "ArrowRight", { shift: true })
    press(c, "Enter")
    expect(active(c)).toBe("A2")
    press(c, "Enter")
    expect(active(c)).toBe("B1")
    expect(rangeName(c.selection.ranges[0])).toBe("A1:B2")
  })

  it("F2 opens the cell's text in edit mode; arrows then stay in the editor", () => {
    const c = setup()
    type(c, "hello")
    press(c, "Enter")
    press(c, "ArrowUp")
    press(c, "F2")
    expect(c.edit?.text).toBe("hello")
    expect(c.mode).toBe("edit")
    expect(press(c, "ArrowLeft")).toBe(false)
    expect(c.edit).not.toBeNull()
  })

  it("in enter mode an arrow commits and moves", () => {
    const c = setup()
    type(c, "7")
    expect(press(c, "ArrowRight")).toBe(true)
    expect(value(c, "A1")).toBe(7)
    expect(active(c)).toBe("B1")
  })

  it("Delete clears contents and undo restores them with the selection", () => {
    const c = setup()
    type(c, "5")
    press(c, "Enter")
    press(c, "ArrowUp")
    press(c, "Delete")
    expect(value(c, "A1")).toBeUndefined()
    press(c, "ArrowDown")
    press(c, "ArrowDown")
    press(c, "z", { primary: true })
    expect(value(c, "A1")).toBe(5)
    expect(active(c)).toBe("A1")
    press(c, "y", { primary: true })
    expect(value(c, "A1")).toBeUndefined()
  })
})

describe("clipboard", () => {
  function withData() {
    const c = setup()
    for (const row of [
      ["1", "a"],
      ["2", "b"],
    ]) {
      type(c, row[0])
      press(c, "Tab")
      type(c, row[1])
      press(c, "Enter")
    }
    c.goTo({ r0: 0, c0: 0, r1: 1, c1: 1 })
    return c
  }

  it("copies as TSV and HTML and pastes its own copy", () => {
    const c = withData()
    const content = c.copy(false)!
    expect(content.text).toBe("1\ta\r\n2\tb\r\n")
    expect(content.html).toContain("<table")
    expect(c.clip).not.toBeNull()
    c.goTo({ r0: 4, c0: 3, r1: 4, c1: 3 })
    c.paste(content.text, content.html)
    expect([value(c, "D5"), value(c, "E5"), value(c, "D6"), value(c, "E6")]).toEqual([1, "a", 2, "b"])
    expect(rangeName(c.selection.ranges[0])).toBe("D5:E6")
    expect(c.clip).not.toBeNull()
  })

  it("a cut pastes as a move and ends the marching ants", () => {
    const c = withData()
    const content = c.copy(true)!
    c.goTo({ r0: 9, c0: 0, r1: 9, c1: 0 })
    c.paste(content.text)
    expect(value(c, "A1")).toBeUndefined()
    expect(value(c, "A10")).toBe(1)
    expect(value(c, "B11")).toBe("b")
    expect(c.clip).toBeNull()
    c.undo()
    expect(value(c, "A1")).toBe(1)
    expect(value(c, "A10")).toBeUndefined()
  })

  it("parses outside text as typed input and tiles over a larger multiple target", () => {
    const c = setup()
    c.goTo({ r0: 0, c0: 0, r1: 3, c1: 0 })
    c.paste("10%\r\nx\r\n")
    expect(value(c, "A1")).toBe(0.1)
    expect(value(c, "A2")).toBe("x")
    expect(value(c, "A3")).toBe(0.1)
    expect(value(c, "A4")).toBe("x")
  })

  it("reads HTML tables when there is no text", () => {
    const c = setup()
    c.paste("", "<table><tr><td>3</td><td>q</td></tr></table>")
    expect([value(c, "A1"), value(c, "B1")]).toEqual([3, "q"])
  })

  it("typing a new entry clears the copy marquee", () => {
    const c = withData()
    c.copy(false)
    c.goTo({ r0: 5, c0: 5, r1: 5, c1: 5 })
    type(c, "z")
    press(c, "Enter")
    expect(c.clip).toBeNull()
  })
})

describe("fill and structure", () => {
  it("fill handle extends a series and Ctrl+D copies down", () => {
    const c = setup()
    type(c, "Mon")
    press(c, "Enter")
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.fillTo({ r0: 0, c0: 0, r1: 3, c1: 0 }, false)
    expect([value(c, "A2"), value(c, "A3"), value(c, "A4")]).toEqual(["Tue", "Wed", "Thu"])
    c.goTo({ r0: 0, c0: 0, r1: 2, c1: 0 })
    press(c, "d", { primary: true })
    expect([value(c, "A2"), value(c, "A3")]).toEqual(["Mon", "Mon"])
  })

  it("inserts and deletes rows with undo", () => {
    const c = setup()
    type(c, "top")
    press(c, "Enter")
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.selectRowsOfSelection()
    c.insertLines("rows")
    expect(value(c, "A2")).toBe("top")
    c.undo()
    expect(value(c, "A1")).toBe("top")
    c.deleteLines("rows")
    expect(value(c, "A1")).toBeUndefined()
  })

  it("Ctrl+9 hides rows and Ctrl+Shift+9 brings them back", () => {
    const c = setup()
    c.goTo({ r0: 2, c0: 0, r1: 3, c1: 0 })
    press(c, "9", { primary: true })
    expect(c.sheet.rows.isHidden(2)).toBe(true)
    expect(c.sheet.rows.isHidden(3)).toBe(true)
    c.goTo({ r0: 1, c0: 0, r1: 4, c1: 0 })
    press(c, "(", { primary: true, shift: true })
    expect(c.sheet.rows.isHidden(2)).toBe(false)
  })

  it("Ctrl+Down jumps to the data edge and the Name Box shows the dragged size", () => {
    const c = setup()
    for (let i = 0; i < 5; i++) {
      type(c, String(i))
      press(c, "Enter")
    }
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    press(c, "ArrowDown", { primary: true })
    expect(active(c)).toBe("A5")
    press(c, "ArrowDown", { primary: true })
    expect(active(c)).toBe("A1048576")
    c.goTo({ r0: 0, c0: 0, r1: 2, c1: 1 })
    expect(c.nameBoxText(true)).toBe("3R x 2C")
    expect(c.nameBoxText(false)).toBe("A1")
  })

  it("End mode makes the next arrow a data-edge jump", () => {
    const c = setup()
    type(c, "1")
    press(c, "Tab")
    type(c, "2")
    press(c, "Enter")
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    press(c, "End")
    expect(c.endMode).toBe(true)
    press(c, "ArrowRight")
    expect(active(c)).toBe("B1")
    expect(c.endMode).toBe(false)
  })
})

describe("typing formulas", () => {
  const typeAt = (c: SheetController, text: string) => {
    if (!c.edit) c.beginEdit("enter", "")
    c.setEditText((c.edit?.text ?? "") + text)
  }

  it("points at cells with the arrow keys, Shift extends, and the entry is calculated", () => {
    const c = setup()
    type(c, "5")
    press(c, "Enter")
    type(c, "7")
    press(c, "Enter") // A3 is active
    typeAt(c, "=")
    expect(c.canPoint).toBe(true)
    press(c, "ArrowUp")
    press(c, "ArrowUp")
    expect(c.edit?.text).toBe("=A1")
    expect(c.mode).toBe("point")
    press(c, "ArrowDown", { shift: true })
    expect(c.edit?.text).toBe("=A1:A2")
    typeAt(c, "+")
    expect(c.mode).toBe("enter")
    press(c, "ArrowUp")
    expect(c.edit?.text).toBe("=A1:A2+A2")
    press(c, "Enter")
    expect(c.sheet.cells.input(2, 0)).toBe("=A1:A2+A2")
    expect(active(c)).toBe("A4")
  })

  it("arrow keys commit when the caret is not where a reference can go", () => {
    const c = setup()
    typeAt(c, "=5")
    press(c, "ArrowDown")
    expect(c.edit).toBeNull()
    expect(value(c, "A1")).toBe(5)
  })

  it("clicking writes a reference (pointAt) and F4 cycles its $ signs", () => {
    const c = setup()
    typeAt(c, "=SUM(")
    c.pointAt({ row: 1, col: 1 }, { row: 3, col: 2 })
    expect(c.edit?.text).toBe("=SUM(B2:C4")
    expect(c.editReferences()).toMatchObject([{ range: { r0: 1, c0: 1, r1: 3, c1: 2 }, pointing: true }])
    press(c, "F4")
    expect(c.edit?.text).toBe("=SUM($B$2:$C$4")
    press(c, "Enter")
    // The missing parenthesis is added on entry.
    expect(c.sheet.cells.input(0, 0)).toBe("=SUM($B$2:$C$4)")
  })

  it("offers functions as their names are typed; Tab inserts one", () => {
    const c = setup()
    typeAt(c, "=su")
    const list = c.autocomplete!
    expect(list.items[0].name).toBe("SUBSTITUTE")
    press(c, "ArrowDown")
    expect(c.autocomplete!.items[c.autocomplete!.pick].name).toBe("SUBTOTAL")
    press(c, "ArrowDown")
    press(c, "Tab")
    expect(c.edit?.text).toBe("=SUM(")
    expect(c.autocomplete).toBeNull()
  })

  it("Ctrl+` shows formulas, F9 calculates in manual mode, Ctrl+Shift+Enter stays put", () => {
    const c = setup()
    typeAt(c, "=1+1")
    press(c, "Enter", { primary: true, shift: true })
    expect(active(c)).toBe("A1")
    expect(value(c, "A1")).toBe(2)
    press(c, "`", { primary: true })
    expect(c.showFormulas).toBe(true)
    c.wb.setCalcMode("manual")
    typeAt(c, "=NOW()*0+3")
    press(c, "Enter")
    c.selectCellAt({ row: 1, col: 1 })
    typeAt(c, "=A1*2")
    press(c, "Enter")
    expect(value(c, "B2")).toBe(6)
    press(c, "F9")
    expect(c.wb.calcStatus.pending).toBe(false)
  })
})

describe("live zoom (a pinch)", () => {
  it("tells only the grid while pinching, and everyone otherwise", () => {
    const c = new SheetController()
    let all = 0
    let scroll = 0
    c.subscribe(() => all++)
    c.subscribeScroll(() => scroll++)
    c.setZoom(1.5, true)
    expect(c.zoom).toBe(1.5)
    expect([all, scroll]).toEqual([0, 1])
    c.setZoom(2)
    expect(c.zoom).toBe(2)
    expect([all, scroll]).toEqual([1, 1])
  })
})
