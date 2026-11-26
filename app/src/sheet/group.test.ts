import { describe, expect, it } from "vitest"
import { MAX_COLS } from "@cells/core"
import { SheetController, type KeyInput } from "./controller"
import { GROUP_MESSAGE } from "./group"

function setup(sheets = 3) {
  const c = new SheetController()
  c.viewWidth = 1000
  c.viewHeight = 600
  for (let i = 1; i < sheets; i++) c.sheets.insert(i)
  c.activate(0)
  c.wb.clearHistory?.()
  return c
}

const press = (c: SheetController, key: string, mods: Partial<Omit<KeyInput, "key">> = {}) => {
  const input = { key, primary: false, shift: false, alt: false, ...mods }
  return c.edit ? c.keyEditing(input) : c.keyReady(input)
}

const cell = (c: SheetController, sheet: number, row: number, col: number) => c.wb.sheets[sheet].cells.value(row, col)

describe("grouped sheets", () => {
  it("Shift+click groups a run of sheets, Ctrl+click adds or takes one out", () => {
    const c = setup(4)
    c.group.click(2, { ctrl: false, shift: true })
    expect(c.group.indexes()).toEqual([0, 1, 2])
    c.group.click(1, { ctrl: true, shift: false })
    expect(c.group.indexes()).toEqual([0, 2])
    c.group.click(3, { ctrl: true, shift: false })
    expect(c.group.indexes()).toEqual([0, 2, 3])
    expect(c.group.active).toBe(true)
  })

  it("typing goes into the same cell of every grouped sheet, as one undo step", () => {
    const c = setup()
    c.group.selectAll()
    c.beginEdit("enter", "42")
    press(c, "Enter")
    expect([cell(c, 0, 0, 0), cell(c, 1, 0, 0), cell(c, 2, 0, 0)]).toEqual([42, 42, 42])
    c.undo()
    expect([cell(c, 0, 0, 0), cell(c, 1, 0, 0), cell(c, 2, 0, 0)]).toEqual([undefined, undefined, undefined])
    // Undo kept the group.
    expect(c.group.active).toBe(true)
  })

  it("formats repeat on the group; a sheet outside it isn't touched", () => {
    const c = setup()
    c.group.click(1, { ctrl: true, shift: false })
    c.format.toggle("bold")
    const bold = (i: number) => c.wb.styles.get(c.wb.sheets[i].cells.get(0, 0)?.style ?? 0).bold
    expect([bold(0), bold(1), bold(2)]).toEqual([true, true, undefined])
  })

  it("Delete clears what each grouped sheet has in the selection", () => {
    const c = setup(2)
    c.wb.transact([{ kind: "setCells", sheet: 1, cells: [{ row: 0, col: 0, data: { value: 7 } }] }], {
      label: "x",
      before: c.selection,
    })
    c.group.selectAll()
    c.clearContents()
    expect(cell(c, 1, 0, 0)).toBeUndefined()
  })

  it("a click on a tab outside a partial group ungroups; inside it the group stays", () => {
    const c = setup(4)
    c.group.set([0, 1])
    c.group.click(1, { ctrl: false, shift: false })
    expect(c.sheetIndex).toBe(1)
    expect(c.group.indexes()).toEqual([0, 1])
    c.group.click(3, { ctrl: false, shift: false })
    expect(c.group.active).toBe(false)
  })

  it("with every sheet grouped, a click on another tab ungroups (Excel)", () => {
    const c = setup()
    c.group.selectAll()
    c.group.click(2, { ctrl: false, shift: false })
    expect(c.sheetIndex).toBe(2)
    expect(c.group.active).toBe(false)
  })

  it("Ctrl+Shift+PgDn adds the next sheet and shows it", () => {
    const c = setup()
    press(c, "PageDown", { primary: true, shift: true })
    expect(c.sheetIndex).toBe(1)
    expect(c.group.indexes()).toEqual([0, 1])
  })

  it("one-sheet features are refused while grouped", () => {
    const c = setup(2)
    let said = ""
    c.shell.alert = (m) => void (said = m)
    c.group.selectAll()
    c.transact([{ kind: "validations", sheet: 0, entries: [] }], { label: "x", before: c.selection })
    expect(said).toBe(GROUP_MESSAGE)
  })

  it("Fill Across Worksheets copies contents, formats or both", () => {
    const c = setup(2)
    c.beginEdit("enter", "5")
    press(c, "Enter")
    c.selection = { ...c.selection, ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], active: { row: 0, col: 0 } }
    c.format.toggle("bold")
    c.group.selectAll()
    c.group.fillAcross("contents")
    expect(cell(c, 1, 0, 0)).toBe(5)
    expect(c.wb.sheets[1].cells.get(0, 0)?.style ?? 0).toBe(0)
    c.group.fillAcross("formats")
    expect(c.wb.sheets[1].cells.get(0, 0)?.style).toBe(c.wb.sheets[0].cells.get(0, 0)?.style)
  })

  it("Fill Across Worksheets on a whole row copies its format and height too (Excel 365)", () => {
    const c = setup(2)
    c.beginEdit("enter", "x")
    press(c, "Enter")
    c.selection = { ...c.selection, ranges: [{ r0: 0, c0: 0, r1: 0, c1: MAX_COLS - 1 }], active: { row: 0, col: 0 } }
    c.format.toggle("bold")
    c.setLineSize("rows", [0], 40)
    c.group.selectAll()
    c.group.fillAcross("all")
    const [one, two] = c.wb.sheets
    expect(two.cells.value(0, 0)).toBe("x")
    expect(two.rowStyles.get(0)).toBe(one.rowStyles.get(0))
    expect(two.rowStyles.get(0)).not.toBe(0)
    expect(two.rows.storedSize(0)).toBe(40)
    c.undo()
    expect(two.rowStyles.get(0)).toBe(0)
    expect(two.rows.hasCustomSize(0)).toBe(false)
  })

  it("grouped sheets are saved as selected and open grouped", () => {
    const c = setup()
    c.group.set([0, 2])
    const views = c.saveViews()
    expect(views.map((v) => v.selected)).toEqual([true, false, true])
    const d = new SheetController(c.wb)
    d.loadViews(views, 0)
    expect(d.group.indexes()).toEqual([0, 2])
  })

  it("Delete on a group removes every grouped sheet", () => {
    const c = setup(4)
    c.activate(1)
    c.group.set([1, 2])
    c.sheets.remove(1)
    expect(c.wb.sheets.length).toBe(2)
    expect(c.group.active).toBe(false)
  })
})
