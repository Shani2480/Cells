import { describe, expect, it, vi } from "vitest"
import { DEFAULT_VALIDATION, type Validation } from "@cells/core"
import { SheetController } from "./controller"
import type { AlertAnswer } from "./validation"
import { fieldFormula, fieldLabels, fieldText, MIXED_RULES, PARTLY_RULED } from "./validation"

const key = (c: SheetController, k: string, o: Partial<{ shift: boolean; alt: boolean; primary: boolean }> = {}) =>
  c.keyEditing({ key: k, primary: false, shift: false, alt: false, ...o })

const whole: Validation = { ...DEFAULT_VALIDATION, type: "whole", formula1: "1", formula2: "10" }

/** A controller with A1:A10 limited to whole numbers 1–10, and an alert that answers as told. */
function ruled(style: Validation["errorStyle"] = "stop") {
  const c = new SheetController()
  c.goTo({ r0: 0, c0: 0, r1: 9, c1: 0 })
  c.validation.apply([{ r0: 0, c0: 0, r1: 9, c1: 0 }], { ...whole, errorStyle: style, error: "1 to 10 only" })
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  const answers: AlertAnswer[] = []
  const alert = vi.fn((_a, answer: (a: AlertAnswer) => void) => answer(answers.shift() ?? "cancel"))
  c.shell.validationAlert = alert
  return { c, answers, alert }
}

describe("Data Validation on entry", () => {
  it("Stop: Retry keeps the entry in the cell to fix, Cancel throws it away", () => {
    const { c, answers, alert } = ruled()
    c.beginEdit("enter", "12")
    answers.push("retry")
    key(c, "Enter")
    expect(alert.mock.calls[0][0]).toEqual({ style: "stop", title: "Cells", message: "1 to 10 only" })
    expect(c.edit?.text).toBe("12")
    expect(c.edit && [c.edit.selStart, c.edit.selEnd]).toEqual([0, 2])
    expect(c.sheet.cells.get(0, 0)).toBeUndefined()
    expect(c.selection.active).toEqual({ row: 0, col: 0 })
    answers.push("cancel")
    key(c, "Enter")
    expect(c.edit).toBeNull()
    expect(c.sheet.cells.get(0, 0)).toBeUndefined()
  })

  it("a valid entry goes in and Enter moves on", () => {
    const { c, alert } = ruled()
    c.beginEdit("enter", "7")
    key(c, "Enter")
    expect(alert).not.toHaveBeenCalled()
    expect(c.sheet.cells.get(0, 0)?.value).toBe(7)
    expect(c.selection.active).toEqual({ row: 1, col: 0 })
  })

  it("Warning: Yes enters it and moves on; No goes back to the entry", () => {
    const { c, answers } = ruled("warning")
    c.beginEdit("enter", "99")
    answers.push("no")
    key(c, "Tab")
    expect(c.edit?.text).toBe("99")
    answers.push("yes")
    key(c, "Tab")
    expect(c.sheet.cells.get(0, 0)?.value).toBe(99)
    expect(c.selection.active).toEqual({ row: 0, col: 1 })
  })

  it("Information: OK enters it; a formula is checked by its result", () => {
    const { c, answers, alert } = ruled("information")
    c.beginEdit("enter", "=5*3")
    answers.push("ok")
    key(c, "Enter")
    expect(alert).toHaveBeenCalledTimes(1)
    expect(c.sheet.cells.get(0, 0)?.input).toBe("=5*3")
  })

  it("a picked list item goes in as typed, and undoes in one step", () => {
    const c = new SheetController()
    const list = { ...DEFAULT_VALIDATION, type: "list" as const, formula1: '"Apple,Banana,Cherry"' }
    c.validation.apply([{ r0: 0, c0: 0, r1: 0, c1: 0 }], list)
    expect(c.validation.dropdownAt(0, 0)).toEqual(["Apple", "Banana", "Cherry"])
    expect(c.validation.dropdownAt(1, 0)).toBeNull()
    c.validation.pick(0, 0, "Banana")
    expect(c.sheet.cells.get(0, 0)?.value).toBe("Banana")
    c.undo()
    expect(c.sheet.cells.get(0, 0)).toBeUndefined()
  })
})

describe("Data Validation dialog flow", () => {
  it("opens on the selection's rule, asks before mixing, and Any value clears", () => {
    const { c } = ruled()
    const opened = vi.fn()
    c.shell.dataValidation = opened
    c.shell.confirm = vi.fn((_m, ok) => ok())
    c.goTo({ r0: 2, c0: 0, r1: 3, c1: 0 })
    c.validation.open()
    expect(opened.mock.calls[0][0].rule.type).toBe("whole")
    // Part of the selection without a rule: Excel offers to extend it.
    c.goTo({ r0: 8, c0: 0, r1: 12, c1: 0 })
    c.validation.open()
    expect(c.shell.confirm).toHaveBeenLastCalledWith(PARTLY_RULED, expect.any(Function))
    expect(opened.mock.calls[1][0].rule.type).toBe("whole")
    c.validation.apply([{ r0: 0, c0: 1, r1: 0, c1: 1 }], { ...DEFAULT_VALIDATION, type: "decimal", formula1: "0" })
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 1 })
    c.validation.open()
    expect(c.shell.confirm).toHaveBeenLastCalledWith(MIXED_RULES, expect.any(Function))
    expect(opened.mock.calls[2][0].rule).toEqual(DEFAULT_VALIDATION)
    c.validation.apply([{ r0: 0, c0: 0, r1: 0, c1: 1 }], { ...DEFAULT_VALIDATION })
    expect(c.validation.at(0, 0)).toBeNull()
    expect(c.validation.at(1, 0)?.rule.type).toBe("whole")
  })

  it("Apply to all other cells with the same settings reaches every cell of the old rule", () => {
    const { c } = ruled()
    c.validation.apply([{ r0: 0, c0: 3, r1: 0, c1: 3 }], whole)
    const changed = { ...whole, formula2: "20" }
    c.validation.apply([{ r0: 0, c0: 0, r1: 0, c1: 0 }], changed, {
      ...whole,
      errorStyle: "stop",
      error: "1 to 10 only",
    })
    expect(c.validation.at(5, 0)?.rule.formula2).toBe("20")
    expect(c.validation.at(0, 3)?.rule.formula2).toBe("10")
  })

  it("Circle Invalid Data circles the cells that break their rule; fixing one drops its circle", () => {
    const { c } = ruled()
    c.sheet.cells.set(0, 0, { value: 50 })
    c.sheet.cells.set(1, 0, { value: 5 })
    c.sheet.cells.set(2, 0, { value: "x" })
    c.validation.circleInvalid()
    expect(c.validation.circled().map((r) => r.r0)).toEqual([0, 2])
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.beginEdit("enter", "3")
    key(c, "Enter")
    expect(c.validation.circled().map((r) => r.r0)).toEqual([2])
    c.validation.clearCircles()
    expect(c.validation.circled()).toEqual([])
  })
})

describe("the dialog's fields", () => {
  it("show and read formulas, dates, times and typed lists as Excel does", () => {
    expect(fieldText('"Apple,Banana"', "list")).toBe("Apple,Banana")
    expect(fieldText("Lists!$A$1:$A$3", "list")).toBe("=Lists!$A$1:$A$3")
    expect(fieldText("46023", "date")).toBe("1/1/2026")
    expect(fieldText("0.5", "time")).toBe("12:00:00 PM")
    expect(fieldText("DATE(2026,1,1)", "date")).toBe("=DATE(2026,1,1)")
    expect(fieldFormula("1/1/2026", "date")).toBe("46023")
    expect(fieldFormula("12:00 PM", "time")).toBe("0.5")
    expect(fieldFormula('Yes,"No"', "list")).toBe('"Yes,""No"""')
    expect(fieldFormula("=$H$1:$H$3", "list")).toBe("$H$1:$H$3")
    expect(fieldFormula("abc", "whole")).toBeNull()
    expect(fieldFormula("", "whole")).toBeNull()
    expect(fieldLabels("date", "greaterThan")).toEqual(["Start date", null])
    expect(fieldLabels("whole", "between")).toEqual(["Minimum", "Maximum"])
    expect(fieldLabels("textLength", "equal")).toEqual(["Length", null])
  })
})
