import { describe, expect, it } from "vitest"
import { normalizeFormulaText, parseFormula, printFormula, tokenize } from "../src/formula"

const roundTrip = (source: string) => {
  const result = parseFormula(source)
  if (!result.ok) throw new Error(`${source}: ${result.message}`)
  return printFormula(result.ast)
}

describe("parser golden cases", () => {
  const cases: [string, string][] = [
    ["1+2*3", "1+2*3"],
    ["(1+2)*3", "(1+2)*3"],
    ["-2^2", "-2^2"],
    ["-(2^2)", "-(2^2)"],
    ["2^-2", "2^-2"],
    ["2^3^2", "2^3^2"],
    ["2^(3^2)", "2^(3^2)"],
    ["10%", "10%"],
    ["-10%^2", "-10%^2"],
    ["1-2-3", "1-2-3"],
    ["1-(2-3)", "1-(2-3)"],
    ['"a"&"b"&1+2', '"a"&"b"&1+2'],
    ["1+2=3", "1+2=3"],
    ["a1+$b$2+c$3+$d4", "A1+$B$2+C$3+$D4"],
    ["SUM(A1:B2)", "SUM(A1:B2)"],
    ["sum(a:a,3:3,$A:$C,$1:$5)", "SUM(A:A,3:3,$A:$C,$1:$5)"],
    ["Sheet2!A1", "Sheet2!A1"],
    ["'My Sheet'!A1:B2", "'My Sheet'!A1:B2"],
    ["'It''s'!A1", "'It''s'!A1"],
    ["Sheet1:Sheet3!A1", "Sheet1:Sheet3!A1"],
    ["SUM('Jan 1:Mar 1'!B2)", "SUM('Jan 1:Mar 1'!B2)"],
    ["{1,2;3,4}", "{1,2;3,4}"],
    ['{1,-2.5,"x";TRUE,#N/A,0}', '{1,-2.5,"x";TRUE,#N/A,0}'],
    ["IF(A1,,2)", "IF(A1,,2)"],
    ["IF(A1,)", "IF(A1,)"],
    ["NOW()", "NOW()"],
    ["A1:B2 B1:C3", "A1:B2 B1:C3"],
    ["SUM((A1,B1:B3))", "SUM((A1,B1:B3))"],
    ["A1#", "A1#"],
    ["SUM(A1#)", "SUM(A1#)"],
    ["@A1:A10", "@A1:A10"],
    ["A1:INDEX(B:B,3)", "A1:INDEX(B:B,3)"],
    ["MyName*2", "MyName*2"],
    ["Sheet1!MyName", "Sheet1!MyName"],
    ["_xlfn.XLOOKUP(1,A:A,B:B)", "XLOOKUP(1,A:A,B:B)"],
    ["1E+3+2.5e-2", "1000+0.025"],
    ['"say ""hi"""', '"say ""hi"""'],
    ["true+FALSE", "TRUE+FALSE"],
    ["#div/0!", "#DIV/0!"],
    ["LOG10(100)", "LOG10(100)"],
    ["XFD1048576", "XFD1048576"],
    ["Sheet1!#REF!+1", "Sheet1!#REF!+1"],
    ["Table1[Amount]", "Table1[Amount]"],
    ["1 + 2", "1+2"],
    ["SUM( A1 , B1 )", "SUM(A1,B1)"],
    ["-A1:B2", "-A1:B2"],
  ]
  it.each(cases)("%s → %s", (source, printed) => {
    expect(roundTrip(source)).toBe(printed)
    // Printing is stable: parsing the printed text prints the same.
    expect(roundTrip(printed)).toBe(printed)
  })

  it("unary minus binds tighter than ^", () => {
    const r = parseFormula("-2^2")
    expect(r.ok && r.ast.type === "binary" && r.ast.op === "^").toBe(true)
  })

  it("out-of-sheet references are names", () => {
    const r = parseFormula("XFE1")
    expect(r.ok && r.ast.type).toBe("name")
    const r2 = parseFormula("A1048577")
    expect(r2.ok && r2.ast.type).toBe("name")
  })

  it("whole rows and columns parse as spans", () => {
    const r = parseFormula("$B:D")
    expect(r.ok && r.ast).toMatchObject({
      type: "ref",
      from: { col: 1, row: null, colAbs: true },
      to: { col: 3, row: null, colAbs: false },
    })
  })
})

describe("parse errors", () => {
  it.each([
    ["", "empty"],
    ["1+", "ends"],
    ["SUM(1,2", "closing"],
    ["(1+2", "closing"],
    ["1+2)", "extra"],
    ['"abc', "quote"],
    ["{1,2;3}", "same length"],
    ["{A1}", "Array constants"],
    ["1 2", "Unexpected"],
  ])("%j fails (%s)", (source, words) => {
    const r = parseFormula(source)
    expect(r.ok).toBe(false)
    if (!r.ok) expect(r.message).toContain(words)
  })
})

describe("text normalization", () => {
  it("upper-cases references and function names but keeps spacing and quoted sheet names", () => {
    expect(normalizeFormulaText(tokenize("sum( a1 , 'my sheet'!b2 )+ true"))).toBe("SUM( A1 , 'my sheet'!B2 )+ TRUE")
  })

  it("tokens cover every character", () => {
    const source = `IF( A1>=10 , "ok" , {1,2;3,4} )&Sheet2!$B$3:C4 % `
    expect(
      tokenize(source)
        .map((t) => t.text)
        .join(""),
    ).toBe(source)
  })
})
