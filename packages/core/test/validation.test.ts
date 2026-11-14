// Data Validation checked against Excel 365 (Range.Validation.Add, then Validation.Value for each typed entry).
import { describe, expect, it } from "vitest"
import {
  applyOp,
  DEFAULT_VALIDATION,
  invalidCells,
  parseInput,
  ValidationChecker,
  validationAt,
  withValidation,
  Workbook,
  type Validation,
  type ValidationOperator,
  type ValidationType,
} from "../src"

function wb() {
  const w = new Workbook()
  const s = w.sheets[0]
  ;["Red", "Green", "Blue"].forEach((v, i) => s.cells.set(i, 7, { value: v }))
  return w
}

/** Which typed entries Excel accepts under the rule (true/false per entry), as Validation.Value reported. */
function check(
  type: ValidationType,
  operator: ValidationOperator,
  f1: string,
  f2: string,
  entries: string[],
  allowBlank = true,
): boolean[] {
  const w = wb()
  const s = w.sheets[0]
  const rule: Validation = { ...DEFAULT_VALIDATION, type, operator, formula1: f1, formula2: f2, allowBlank }
  s.validations = withValidation([], [{ r0: 0, c0: 0, r1: 0, c1: 0 }], rule)
  const checker = new ValidationChecker(w.sheets, 0, 46000)
  return entries.map((text) => {
    const value =
      text === ""
        ? undefined
        : text === "=NA()"
          ? { error: "#N/A" as const }
          : text === "=5"
            ? 5
            : parseInput(text)?.value
    return checker.valid(s.validations[0], 0, 0, value)
  })
}

describe("Data Validation: which values pass (as Excel)", () => {
  it("whole numbers between 1 and 10", () => {
    expect(
      check("whole", "between", "1", "10", ["5", "5.5", "0", "10", "'5", "", "TRUE", "abc", "=5", "=NA()", "1E1"]),
    ).toEqual([true, false, false, true, false, true, false, false, true, false, true])
    expect(check("whole", "between", "1", "10", [""], false)).toEqual([false])
  })
  it("decimals, dates, times and text length", () => {
    expect(check("decimal", "greaterThan", "1.5", "", ["2", "1.5", "1.6", "'2", "abc"])).toEqual([
      true,
      false,
      true,
      false,
      false,
    ])
    expect(check("date", "between", "46082", "46112", ["3/15/2026", "4/1/2026", "46085", "abc"])).toEqual([
      true,
      false,
      true,
      false,
    ])
    expect(check("time", "lessThan", "0.5", "", ["9:30", "13:00", "0.4", "1.2"])).toEqual([true, false, true, false])
    expect(check("textLength", "lessThanOrEqual", "3", "", ["abc", "abcd", "123", "1234", "12.5"])).toEqual([
      true,
      false,
      true,
      false,
      false,
    ])
  })
  it("a typed list matches case and all; a range list ignores case", () => {
    expect(check("list", "between", '"a,b, c"', "", ["a", "A", "b", " c", "c", "d", "1"])).toEqual([
      true,
      false,
      true,
      true,
      true,
      false,
      false,
    ])
    expect(check("list", "between", '"1,2,3"', "", ["1", "'1", "2.0", "4"])).toEqual([true, true, true, false])
    expect(check("list", "between", "$H$1:$H$3", "", ["Red", "red", "Purple"])).toEqual([true, true, false])
  })
  it("custom formulas, and a bound in an empty cell counts as 0", () => {
    expect(check("custom", "between", "ISNUMBER(A1)", "", ["5", "x", ""])).toEqual([true, false, true])
    expect(check("custom", "between", "A1>3", "", ["", "5", "2"], false)).toEqual([false, true, false])
    expect(check("whole", "greaterThan", "$J$1", "", ["3", "7"])).toEqual([true, true])
  })
})

describe("Data Validation rules on a sheet", () => {
  it("a cell has one rule; setting one cuts the old rule around it; undo puts it back", () => {
    const w = wb()
    const s = w.sheets[0]
    const whole = { ...DEFAULT_VALIDATION, type: "whole" as const, formula1: "1", formula2: "10" }
    s.validations = withValidation([], [{ r0: 0, c0: 0, r1: 9, c1: 0 }], whole)
    const list = { ...DEFAULT_VALIDATION, type: "list" as const, formula1: '"x,y"' }
    const next = withValidation(s.validations, [{ r0: 4, c0: 0, r1: 4, c1: 0 }], list)
    const undo = applyOp(w.sheets, { kind: "validations", sheet: 0, entries: next })
    expect(validationAt(s, 4, 0)?.rule.type).toBe("list")
    expect(validationAt(s, 3, 0)?.rule.type).toBe("whole")
    expect(validationAt(s, 5, 0)?.rule.type).toBe("whole")
    for (const op of undo.inverse) applyOp(w.sheets, op)
    expect(validationAt(s, 4, 0)?.rule.type).toBe("whole")
  })

  it("follows inserted rows, and Circle Invalid Data finds the cells that break their rule", () => {
    const w = wb()
    const s = w.sheets[0]
    s.validations = withValidation([], [{ r0: 0, c0: 0, r1: 3, c1: 0 }], {
      ...DEFAULT_VALIDATION,
      type: "whole",
      formula1: "1",
      formula2: "10",
    })
    applyOp(w.sheets, { kind: "insertRows", sheet: 0, at: 0, count: 2 })
    expect(s.validations[0].ranges).toEqual([{ r0: 2, c0: 0, r1: 5, c1: 0 }])
    s.cells.set(2, 0, { value: 5 })
    s.cells.set(3, 0, { value: 50 })
    s.cells.set(4, 0, { value: "x" })
    expect(invalidCells(w.sheets, 0, 46000)).toEqual([
      { r0: 3, c0: 0, r1: 3, c1: 0 },
      { r0: 4, c0: 0, r1: 4, c1: 0 },
    ])
  })
})

describe("checking an entry before it goes in", () => {
  it("works out formulas and skips rules without an error alert", () => {
    const w = wb()
    const s = w.sheets[0]
    s.cells.set(0, 2, { value: 7 })
    const whole: Validation = { ...DEFAULT_VALIDATION, type: "whole", formula1: "1", formula2: "10" }
    s.validations = withValidation([], [{ r0: 0, c0: 0, r1: 4, c1: 0 }], whole)
    s.validations = withValidation(s.validations, [{ r0: 0, c0: 1, r1: 0, c1: 1 }], { ...whole, showError: false })
    const checker = new ValidationChecker(w.sheets, 0, 46000)
    expect(checker.rejects("5", 0, 0)).toBeNull()
    expect(checker.rejects("11", 0, 0)?.rule.type).toBe("whole")
    expect(checker.rejects("=C1+1", 0, 0)).toBeNull()
    expect(checker.rejects("=C1*2", 0, 0)).not.toBeNull()
    expect(checker.rejects("", 0, 0)).toBeNull()
    expect(checker.rejects("abc", 0, 1)).toBeNull()
    expect(checker.rejects("abc", 0, 5)).toBeNull()
  })
})
