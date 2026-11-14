import { describe, expect, it } from "vitest"
import {
  canPointAt,
  completeFormula,
  editReferences,
  FUNCTION_INFO,
  FUNCTIONS,
  functionAtCaret,
  functionsStartingWith,
  nameAtCaret,
  pointedText,
  toggleAbsolute,
} from "../src/formula"

describe("formula editing helpers", () => {
  it("finds references with stable colours", () => {
    const refs = editReferences("=A1+SUM(B2:C3)*a1+Sheet2!D4")
    expect(refs.map((r) => [r.start, r.end, r.color, r.sheet])).toEqual([
      [1, 3, 0, null],
      [8, 13, 1, null],
      [15, 17, 0, null],
      [18, 27, 2, "Sheet2"],
    ])
    expect(refs[1].range).toEqual({ r0: 1, c0: 1, r1: 2, c1: 2 })
    expect(editReferences("hello")).toEqual([])
  })

  it.each([
    ["=", 1, true],
    ["=A1+", 4, true],
    ["=SUM(", 5, true],
    ["=SUM(A1,", 8, true],
    ["=SUM(A1, ", 9, true],
    ["=A1", 3, false],
    ["=A1", 2, false],
    ["=5%", 3, false],
    ['="x"&', 5, true],
    ["abc", 3, false],
  ])("can point in %j at %i: %s", (text, caret, expected) => {
    expect(canPointAt(text, caret)).toBe(expected)
  })

  it("F4 cycles the reference at the caret", () => {
    let state = { text: "=A1+B2", start: 0, end: 3 }
    const seen: string[] = []
    for (let i = 0; i < 4; i++) {
      state = toggleAbsolute(state.text, 2)!
      seen.push(state.text)
    }
    expect(seen).toEqual(["=$A$1+B2", "=A$1+B2", "=$A1+B2", "=A1+B2"])
    expect(toggleAbsolute("=SUM(A1:B2)", 9)).toEqual({ text: "=SUM($A$1:$B$2)", start: 5, end: 14 })
    expect(toggleAbsolute("=Sheet2!C3", 10)!.text).toBe("=Sheet2!$C$3")
    expect(toggleAbsolute("=A:A", 4)!.text).toBe("=$A:$A")
    expect(toggleAbsolute("=1+2", 2)).toBeNull()
  })

  it("knows the function and argument around the caret", () => {
    expect(functionAtCaret("=SUM(", 5)).toMatchObject({ info: { name: "SUM" }, arg: 0 })
    expect(functionAtCaret("=IF(A1>0,ROUND(A1,", 18)).toMatchObject({ info: { name: "ROUND" }, arg: 1 })
    expect(functionAtCaret("=IF(A1>0,ROUND(A1,2),", 21)).toMatchObject({ info: { name: "IF" }, arg: 2 })
    expect(functionAtCaret("=IF((1,2),", 10)).toMatchObject({ info: { name: "IF" }, arg: 1 })
    expect(functionAtCaret("=1+2", 4)).toBeNull()
  })

  it("autocompletes function names", () => {
    expect(nameAtCaret("=SU", 3)).toEqual({ start: 1, prefix: "SU" })
    expect(nameAtCaret("=A1+vlo", 7)).toEqual({ start: 4, prefix: "vlo" })
    expect(nameAtCaret("=A1", 3)).toEqual({ start: 1, prefix: "A1" })
    expect(nameAtCaret("=A1+", 4)).toBeNull()
    expect(functionsStartingWith("sumi").map((f) => f.name)).toEqual(["SUMIF", "SUMIFS"])
  })

  it("writes pointed references", () => {
    expect(pointedText({ r0: 0, c0: 0, r1: 0, c1: 0 })).toBe("A1")
    expect(pointedText({ r0: 1, c0: 1, r1: 4, c1: 2 }, "Sheet2!")).toBe("Sheet2!B2:C5")
    expect(pointedText({ r0: 0, c0: 2, r1: 1_048_575, c1: 2 })).toBe("C:C")
  })

  it("completes formulas on entry", () => {
    expect(completeFormula("=sum(a1:b2")).toBe("=SUM(A1:B2)")
    expect(completeFormula('=if(a1>0,"yes",round(b1,2')).toBe('=IF(A1>0,"yes",ROUND(B1,2))')
    expect(completeFormula("='my sheet'!a1+true")).toBe("='my sheet'!A1+TRUE")
    expect(completeFormula("hello")).toBe("hello")
  })

  it("describes every function the engine knows", () => {
    const missing = [...FUNCTIONS.keys()].filter((name) => !FUNCTION_INFO.has(name))
    expect(missing).toEqual([])
    // And autocomplete offers nothing the engine can't calculate.
    const unknown = [...FUNCTION_INFO.keys()].filter((name) => !FUNCTIONS.has(name))
    expect(unknown).toEqual([])
  })
})
