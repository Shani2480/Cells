import { fireEvent, render, screen } from "@testing-library/react"
import { afterEach, describe, expect, it } from "vitest"
import { DEFAULT_AUTOCORRECT, setAutoCorrect, type AutoCorrectSettings } from "@/proofing/autocorrect"
import { AutoCorrectDialog } from "@/settings/AutoCorrectDialog"
import { SheetController } from "./controller"

/** Types `text` into A1 one character at a time, as the cell's text box reports it. */
function type(c: SheetController, text: string) {
  c.beginEdit("enter", "")
  for (const ch of text) {
    const e = c.edit!
    c.setEditText(e.text.slice(0, e.selStart) + ch + e.text.slice(e.selEnd), e.selStart + 1)
  }
}

describe("AutoCorrect while typing in a cell", () => {
  afterEach(() => setAutoCorrect(DEFAULT_AUTOCORRECT))

  it("corrects each word as it's finished, and the last one when the entry goes in", () => {
    const c = new SheetController()
    type(c, "teh total (c) due monday. thanks teh")
    expect(c.edit!.text).toBe("the total © due Monday. Thanks teh")
    expect(c.edit!.selStart).toBe(c.edit!.text.length)
    c.commitEdit()
    expect(c.sheet.cells.value(0, 0)).toBe("the total © due Monday. Thanks the")
  })

  it("leaves formulas alone, and is off when its options are off", () => {
    const c = new SheetController()
    type(c, '=LEN("teh ")')
    c.commitEdit()
    expect(c.sheet.cells.get(0, 0)?.input).toBe('=LEN("teh ")')
    setAutoCorrect({ ...DEFAULT_AUTOCORRECT, replaceText: false })
    type(c, "teh ")
    expect(c.edit!.text).toBe("teh ")
  })
})

describe("the AutoCorrect dialog", () => {
  it("adds, replaces and deletes entries, and turns options off", () => {
    let result: AutoCorrectSettings | null = null
    render(<AutoCorrectDialog value={DEFAULT_AUTOCORRECT} onCancel={() => undefined} onDone={(s) => (result = s)} />)
    fireEvent.change(screen.getByLabelText("Replace:"), { target: { value: "ytd" } })
    fireEvent.change(screen.getByLabelText("With:"), { target: { value: "year to date" } })
    fireEvent.click(screen.getByRole("button", { name: "Add" }))
    // A built-in entry: picking it fills both boxes, and Delete takes it out.
    fireEvent.click(screen.getByRole("option", { name: "teh, the" }))
    expect(screen.getByLabelText("With:")).toHaveValue("the")
    fireEvent.click(screen.getByRole("button", { name: "Delete" }))
    fireEvent.click(screen.getByLabelText("Capitalize names of days"))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(result).toMatchObject({ added: [["ytd", "year to date"]], removed: ["teh"], dayNames: false })
  })
})
