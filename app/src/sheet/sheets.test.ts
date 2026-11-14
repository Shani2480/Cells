import { parseRange, selectRange, writeInputOp } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { SHEET_MESSAGES } from "./sheets"

function setup(names = ["Sheet1"]) {
  const c = new SheetController()
  // A workbook with more sheets: add them as a user would (no undo history kept).
  for (let i = 1; i < names.length; i++) c.sheets.insert(i)
  names.forEach((name, i) => c.sheets.rename(i, name))
  c.wb.clearHistory()
  c.activate(0)
  const messages: string[] = []
  const confirms: string[] = []
  c.shell = { ...c.shell, alert: (m) => messages.push(m), confirm: (m, ok) => (confirms.push(m), ok()) }
  const select = (ref: string) => c.setSelection(selectRange(parseRange(ref)!))
  const put = (ref: string, text: string) => {
    const r = parseRange(ref)!
    c.transact([writeInputOp(c.sheetIndex, c.sheet, c.wb.styles, r.r0, r.c0, text)], {
      label: "Typing",
      before: c.selection,
      after: c.selection,
    })
  }
  const key = (key: string, shift = false) => c.keyReady({ key, primary: true, shift, alt: false })
  const names_ = () => c.wb.sheets.map((s) => s.name)
  return { c, s: c.sheets, select, put, key, messages, confirms, names: names_ }
}

describe("sheet tabs", () => {
  it("each sheet keeps its own selection and scroll", () => {
    const { c, select } = setup(["Sheet1", "Sheet2"])
    select("C5")
    c.scrollTo(300, 500)
    c.activate(1)
    expect(c.sheet.name).toBe("Sheet2")
    expect(c.selection.active).toEqual({ row: 0, col: 0 })
    expect(c.scrollY).toBe(0)
    select("B2")
    c.activate(0)
    expect(c.selection.active).toEqual({ row: 4, col: 2 })
    expect(c.scrollX).toBe(300)
    c.activate(1)
    expect(c.selection.active).toEqual({ row: 1, col: 1 })
  })

  it("Ctrl+PgDn and Ctrl+PgUp step through the visible sheets without wrapping", () => {
    const { c, s, key } = setup(["A", "B", "C"])
    s.hide(1)
    c.activate(0)
    key("PageDown")
    expect(c.sheet.name).toBe("C")
    key("PageDown")
    expect(c.sheet.name).toBe("C")
    key("PageUp")
    expect(c.sheet.name).toBe("A")
    key("PageUp")
    expect(c.sheet.name).toBe("A")
  })

  it("Shift+F11 inserts a sheet before the active one; undo removes it and shows the sheet from before", () => {
    const { c, names } = setup(["Sheet1", "Sheet2"])
    c.activate(1)
    c.keyReady({ key: "F11", primary: false, shift: true, alt: false })
    expect(names()).toEqual(["Sheet1", "Sheet3", "Sheet2"])
    expect(c.sheet.name).toBe("Sheet3")
    c.undo()
    expect(names()).toEqual(["Sheet1", "Sheet2"])
    expect(c.sheet.name).toBe("Sheet2")
    c.redo()
    expect(c.sheet.name).toBe("Sheet3")
  })

  it("undo of an edit on another sheet goes back to that sheet and its selection", () => {
    const { c, select, put } = setup(["Sheet1", "Sheet2"])
    c.activate(1)
    select("D4")
    put("D4", "7")
    c.activate(0)
    c.undo()
    expect(c.sheet.name).toBe("Sheet2")
    expect(c.selection.active).toEqual({ row: 3, col: 3 })
    expect(c.sheet.cells.value(3, 3)).toBeUndefined()
  })

  it("deleting asks first when the sheet holds data, refuses the last visible sheet, and undo restores it", () => {
    const { c, s, put, messages, confirms, names } = setup(["Sheet1", "Data"])
    s.remove(0) // empty: no question
    expect(confirms).toEqual([])
    expect(names()).toEqual(["Data"])
    expect(c.sheet.name).toBe("Data")
    put("A1", "5")
    s.remove(0)
    expect(messages).toEqual([SHEET_MESSAGES.lastVisible])
    c.undo() // the typing
    c.undo() // the delete
    expect(names()).toEqual(["Sheet1", "Data"])
    expect(c.sheet.name).toBe("Sheet1")
    c.activate(1)
    put("A1", "5")
    s.remove(1)
    expect(confirms).toEqual([SHEET_MESSAGES.deleteData])
    expect(names()).toEqual(["Sheet1"])
  })

  it("rename refuses bad and taken names with Excel's message", () => {
    const { s, names } = setup(["Sheet1", "Sheet2"])
    expect(s.rename(1, "sheet1")).toMatch(/already taken/)
    expect(s.rename(1, "a:b")).toMatch(/invalid/)
    expect(s.rename(1, "Totals")).toBeNull()
    expect(names()).toEqual(["Sheet1", "Totals"])
  })

  it("move and copy place the sheet before another; a copy is named 'Name (2)'", () => {
    const { c, s, put, names } = setup(["A", "B", "C"])
    put("A1", "x")
    s.move(0, 3)
    expect(names()).toEqual(["B", "C", "A"])
    expect(c.sheet.name).toBe("A")
    s.copy(2, 0)
    expect(names()).toEqual(["A (2)", "B", "C", "A"])
    expect(c.sheet.cells.value(0, 0)).toBe("x")
    s.copy(3, 4)
    expect(names().at(-1)).toBe("A (3)")
    c.undo()
    c.undo()
    c.undo()
    expect(names()).toEqual(["A", "B", "C"])
  })

  it("hiding the active sheet shows the next one; the last visible sheet can't be hidden", () => {
    const { c, s, messages } = setup(["A", "B"])
    s.hide(0)
    expect(c.sheet.name).toBe("B")
    s.hide(1)
    expect(messages).toEqual([SHEET_MESSAGES.lastVisible])
    s.unhide(0)
    expect(c.sheet.name).toBe("A")
    expect(s.hidden()).toEqual([])
  })

  it("a cut pasted on another sheet moves the cells there", () => {
    const { c, select, put } = setup(["Sheet1", "Sheet2"])
    put("A1", "42")
    select("A1")
    c.copy(true)
    c.activate(1)
    select("C3")
    c.paste(c.clip!.text)
    expect(c.sheet.cells.value(2, 2)).toBe(42)
    expect(c.wb.sheets[0].cells.value(0, 0)).toBeUndefined()
  })

  it("Protect Workbook refuses structure changes but not edits; the password unlocks it", () => {
    const { c, s, messages, names, put } = setup(["Data", "Notes"])
    s.protectWorkbook("secret")
    expect(c.wb.structure.protection?.password).toBe("DAA7")
    s.insert()
    s.rename(1, "Other")
    s.hide(1)
    expect(names()).toEqual(["Data", "Notes"])
    expect(c.wb.sheets[1].hidden).toBe(false)
    expect(messages).toEqual([
      SHEET_MESSAGES.structureLocked,
      SHEET_MESSAGES.structureLocked,
      SHEET_MESSAGES.structureLocked,
    ])
    // Cells and tab colours may still change.
    put("A1", "42")
    s.setColor(1, "#FF0000")
    expect(c.sheet.cells.value(0, 0)).toBe(42)
    expect(c.wb.sheets[1].tabColor).toBe("#FF0000")
    expect(s.unprotectWorkbook("wrong")).toBe(false)
    expect(s.unprotectWorkbook("secret")).toBe(true)
    s.insert()
    expect(names()).toHaveLength(3)
    // Undo takes back the insert, then the unprotect.
    c.wb.undo()
    c.wb.undo()
    expect(s.isWorkbookProtected()).toBe(true)
  })
})
