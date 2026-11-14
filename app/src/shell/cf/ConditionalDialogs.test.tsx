import { parseCell, parseRange, selectRange, writeInputOp } from "@cells/core"
import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { ManageRulesDialog, NewRuleDialog, QuickRuleDialog, valueFormula } from "./ConditionalDialogs"

function setup(cells: Record<string, string> = { A1: "3", A2: "10", A3: "7" }) {
  const c = new SheetController()
  for (const [ref, text] of Object.entries(cells)) {
    const { row, col } = parseCell(ref)!
    c.wb.transact([writeInputOp(0, c.sheet, c.wb.styles, row, col, text)], { label: "Typing" })
  }
  c.setSelection(selectRange(parseRange("A1:A3")!))
  return c
}

describe("conditional formatting dialogs", () => {
  it("values: formulas, numbers and dates as numbers, other text quoted", () => {
    expect(valueFormula("=$B$1")).toBe("$B$1")
    expect(valueFormula("12.5")).toBe("12.5")
    expect(valueFormula("1/2/2026")).toBe("46024")
    expect(valueFormula('say "hi"')).toBe('"say ""hi"""')
  })

  it("Greater Than previews while open and adds the rule on OK (one undo step)", () => {
    const c = setup()
    let closed = false
    render(<QuickRuleDialog controller={c} kind="greaterThan" onClose={() => (closed = true)} />)
    const box = screen.getByRole("textbox", { name: "Value" })
    // Excel suggests a value from the selection.
    expect((box as HTMLInputElement).value).toBe("7")
    fireEvent.change(box, { target: { value: "5" } })
    expect(c.conditional.preview?.rule).toMatchObject({ kind: "cellIs", operator: "greaterThan", values: ["5"] })
    expect(c.cfAt(1, 0)?.style?.fill).toBe("#FFC7CE")
    expect(c.sheet.conditional).toEqual([])
    fireEvent.change(screen.getByRole("combobox", { name: "Format with" }), { target: { value: "2" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(closed).toBe(true)
    expect(c.conditional.preview).toBeNull()
    expect(c.sheet.conditional[0].rule).toMatchObject({ style: { fill: "#C6EFCE", fontColor: "#006100" } })
    expect(c.cfAt(0, 0)).toBeNull()
    expect(c.cfAt(1, 0)?.style?.fill).toBe("#C6EFCE")
    c.undo()
    expect(c.sheet.conditional).toEqual([])
  })

  it("Cancel leaves nothing behind", () => {
    const c = setup()
    const { unmount } = render(<QuickRuleDialog controller={c} kind="top" onClose={() => undefined} />)
    expect(c.conditional.preview).not.toBeNull()
    fireEvent.click(screen.getByRole("button", { name: "Cancel" }))
    unmount()
    expect(c.conditional.preview).toBeNull()
    expect(c.sheet.conditional).toEqual([])
  })

  it("New Rule › Use a formula", () => {
    const c = setup()
    render(<NewRuleDialog controller={c} start="formula" onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(screen.getByRole("alert").textContent).toBe("Enter a formula.")
    fireEvent.change(screen.getByRole("textbox", { name: "Formula" }), { target: { value: "=A1>5" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.conditional[0].rule).toMatchObject({ kind: "expression", formula: "A1>5" })
    expect(c.cfAt(1, 0)?.style).toBeTruthy()
    expect(c.cfAt(0, 0)).toBeNull()
  })

  it("Manage Rules: reorder, stop if true, change Applies to, delete; OK commits once", () => {
    const c = setup()
    c.conditional.add({ kind: "blanks", blank: false, style: { fill: "#FFC7CE" } })
    c.conditional.add({ kind: "top", bottom: false, percent: false, rank: 1, style: { fill: "#C6EFCE" } })
    render(<ManageRulesDialog controller={c} onClose={() => undefined} />)
    const rows = () => screen.getAllByRole("row").slice(1)
    expect(rows().map((r) => r.textContent?.split("AaBb")[0])).toEqual(["Top 1", "No Blanks"])
    fireEvent.click(rows()[1])
    fireEvent.click(screen.getByRole("button", { name: "Move Up" }))
    expect(rows()[0].textContent).toContain("No Blanks")
    const applies = screen.getAllByRole("textbox", { name: "Applies to" })[0]
    expect((applies as HTMLInputElement).value).toBe("=$A$1:$A$3")
    fireEvent.change(applies, { target: { value: "=$A$1:$A$2" } })
    fireEvent.blur(applies)
    fireEvent.click(screen.getAllByRole("checkbox", { name: "Stop If True" })[0])
    fireEvent.click(rows()[1])
    fireEvent.click(screen.getByRole("button", { name: "Delete Rule" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.conditional).toHaveLength(1)
    expect(c.sheet.conditional[0]).toMatchObject({
      rule: { kind: "blanks" },
      stopIfTrue: true,
      ranges: [parseRange("A1:A2")],
    })
    c.undo()
    expect(c.sheet.conditional.map((e) => e.rule.kind)).toEqual(["top", "blanks"])
  })

  it("Clear Rules from Selected Cells and from the Entire Sheet", () => {
    const c = setup()
    c.conditional.add({ kind: "blanks", blank: false, style: { fill: "#FFC7CE" } })
    c.setSelection(selectRange(parseRange("A2")!))
    c.conditional.clearSelected()
    expect(c.sheet.conditional[0].ranges).toEqual([parseRange("A1"), parseRange("A3")])
    c.conditional.clearSheet()
    expect(c.sheet.conditional).toEqual([])
  })
})
