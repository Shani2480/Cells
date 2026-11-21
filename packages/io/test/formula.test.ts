import { describe, expect, it } from "vitest"
import { formulaFromFile, formulaToFile } from "../src/xlsx/formula"

describe("formula text in files", () => {
  it("strips Excel's prefixes on reading", () => {
    expect(formulaFromFile("_xlfn.XLOOKUP(A1,B:B,C:C)")).toBe("=XLOOKUP(A1,B:B,C:C)")
    expect(formulaFromFile("_xlfn._xlws.FILTER(A1:A9,B1:B9>0)")).toBe("=FILTER(A1:A9,B1:B9>0)")
    expect(formulaFromFile("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")).toBe("=LET(x,1,x+1)")
    expect(formulaFromFile("SUM(A1:A3)")).toBe("=SUM(A1:A3)")
  })

  it("adds them on writing, the way Excel does", () => {
    expect(formulaToFile("=XLOOKUP(A1,B:B,C:C)")).toBe("_xlfn.XLOOKUP(A1,B:B,C:C)")
    expect(formulaToFile("=SORT(FILTER(A1:A9,B1:B9>0))")).toBe("_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))")
    // Excel 2010's statistics names carry the prefix; the older ones they replaced don't.
    expect(formulaToFile("=NORM.S.DIST(1,TRUE)+NORMSDIST(1)")).toBe("_xlfn.NORM.S.DIST(1,TRUE)+NORMSDIST(1)")
    expect(formulaToFile("=T.TEST(A1:A5,B1:B5,2,3)*TTEST(A1:A5,B1:B5,2,3)")).toBe(
      "_xlfn.T.TEST(A1:A5,B1:B5,2,3)*TTEST(A1:A5,B1:B5,2,3)",
    )
    expect(formulaToFile("=LET(x,1,y,2,x+y)")).toBe("_xlfn.LET(_xlpm.x,1,_xlpm.y,2,_xlpm.x+_xlpm.y)")
    expect(formulaToFile('=SUM(A1:A3)&"XLOOKUP("')).toBe('SUM(A1:A3)&"XLOOKUP("')
    expect(formulaToFile("=stdev.s(A1:A3)")).toBe("_xlfn.stdev.s(A1:A3)")
  })

  it("writes LAMBDA as Excel saves it, and reads it back", () => {
    // These are the strings Excel 365 wrote for the same formulas and names.
    const saved: [string, string][] = [
      ["=LET(f,LAMBDA(x,x*2),f(21))", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(21))"],
      ["=LAMBDA(x,x+1)(5)", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(5)"],
      ["=MyDouble(4)", "MyDouble(4)"],
      ["=MAP({1,2},LAMBDA(v,v*10))", "_xlfn.MAP({1,2},_xlfn.LAMBDA(_xlpm.v,_xlpm.v*10))"],
      [
        "=LAMBDA(x,y,IF(ISOMITTED(y),x,y))(1,)",
        "_xlfn.LAMBDA(_xlpm.x,_xlpm.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.y))(1,)",
      ],
      [
        "=REDUCE(0,{1,2},LAMBDA(acc,v,acc+v))",
        "_xlfn.REDUCE(0,{1,2},_xlfn.LAMBDA(_xlpm.acc,_xlpm.v,_xlpm.acc+_xlpm.v))",
      ],
      ["=MAKEARRAY(1,2,LAMBDA(r,c,r+c))", "_xlfn.MAKEARRAY(1,2,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c))"],
      ["=LAMBDA(a,b,a+b)", "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)"],
    ]
    for (const [cells, file] of saved) {
      expect(formulaToFile(cells)).toBe(file)
      expect(formulaFromFile(file)).toBe(cells)
    }
  })

  it("marks built-in functions passed by name, and the newest functions", () => {
    expect(formulaToFile("=BYROW(A1:C3,SUM)")).toBe("_xlfn.BYROW(A1:C3,_xleta.SUM)")
    expect(formulaFromFile("_xlfn.BYROW(A1:C3,_xleta.SUM)")).toBe("=BYROW(A1:C3,SUM)")
    expect(formulaToFile("=GROUPBY(A2:A9,C2:C9,PERCENTOF)")).toBe("_xlfn.GROUPBY(A2:A9,C2:C9,_xleta.PERCENTOF)")
    expect(formulaToFile("=ROWS(TRIMRANGE(A:A))")).toBe("ROWS(_xlfn.TRIMRANGE(A:A))")
  })
})
