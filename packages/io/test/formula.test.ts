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
})
