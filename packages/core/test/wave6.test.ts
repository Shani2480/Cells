import { describe, expect, it } from "vitest"
import { parseInput, Workbook, type Op } from "../src"
import { formatCode } from "../src/formula/functions/cell-info"

/** The value (or spilled rows) of a formula entered in E1, next to a little mock table in A1:B3. */
function run(formula: string, ops: Op[] = []): unknown {
  const wb = new Workbook(["Sheet1", "Notes"])
  const texts: [number, number, string][] = [
    [0, 0, "Pens"],
    [0, 1, "12"],
    [1, 0, "Ink"],
    [1, 1, "-3.5"],
    [2, 0, "Paper"],
    [2, 1, "40"],
    [0, 4, formula],
  ]
  const cells = texts.map(([row, col, text]) => {
    const p = parseInput(text)!
    return { row, col, data: { value: p.value, input: p.input } }
  })
  wb.transact([{ kind: "setCells", sheet: 0, cells }, ...ops], { label: "Load", history: false })
  wb.calculate(true)
  const area = wb.calc.spillAt(0, 0, 4)
  const v = (r: number, c: number) => wb.sheets[0].cells.value(r, c) ?? ""
  if (!area) return v(0, 4)
  const out: unknown[][] = []
  for (let r = area.r0; r <= area.r1; r++) {
    const row: unknown[] = []
    for (let c = area.c0; c <= area.c1; c++) row.push(v(r, c))
    out.push(row)
  }
  return out
}

const err = (e: string) => ({ error: e })

describe("REGEXTEST, REGEXEXTRACT, REGEXREPLACE", () => {
  it("tests a pattern, case-sensitive unless asked", () => {
    expect(run('=REGEXTEST("Order 123","\\d+")')).toBe(true)
    expect(run('=REGEXTEST("abc","B")')).toBe(false)
    expect(run('=REGEXTEST("abc","B",1)')).toBe(true)
    expect(run('=REGEXTEST(A1:A3,"^P")')).toEqual([[true], [false], [true]])
    expect(run('=REGEXTEST("abc","(")')).toEqual(err("#VALUE!"))
  })

  it("extracts the first match, all matches down a column, or the groups across a row", () => {
    expect(run('=REGEXEXTRACT("tel 555-1234 or 555-9876","\\d{3}-\\d{4}")')).toBe("555-1234")
    expect(run('=REGEXEXTRACT("tel 555-1234 or 555-9876","\\d{3}-\\d{4}",1)')).toEqual([["555-1234"], ["555-9876"]])
    expect(run('=REGEXEXTRACT("Alex Morgan","(\\w+) (\\w+)",2)')).toEqual([["Alex", "Morgan"]])
    expect(run('=REGEXEXTRACT("abc","\\d")')).toEqual(err("#N/A"))
    expect(run('=REGEXEXTRACT("abc","b",3)')).toEqual(err("#VALUE!"))
  })

  it("replaces every match, or only the nth (from the end when negative), with $n groups", () => {
    expect(run('=REGEXREPLACE("a1b22c333","\\d+","#")')).toBe("a#b#c#")
    expect(run('=REGEXREPLACE("a1b22c333","\\d+","#",2)')).toBe("a1b#c333")
    expect(run('=REGEXREPLACE("a1b22c333","\\d+","#",-1)')).toBe("a1b22c#")
    expect(run('=REGEXREPLACE("Morgan, Alex","(\\w+), (\\w+)","$2 $1")')).toBe("Alex Morgan")
    expect(run('=REGEXREPLACE("x1y2","(\\d)","<$1>",2)')).toBe("x1y<2>")
    expect(run('=REGEXREPLACE("ABC","b","-",0,1)')).toBe("A-C")
  })
})

describe("CELL", () => {
  it("names the number format with Excel's codes", () => {
    expect(formatCode(undefined)).toBe("G")
    expect(formatCode("0")).toBe("F0")
    expect(formatCode("0.00")).toBe("F2")
    expect(formatCode("#,##0")).toBe(",0")
    expect(formatCode("#,##0.00")).toBe(",2")
    expect(formatCode("$#,##0_);($#,##0)")).toBe("C0")
    expect(formatCode("$#,##0.00_);[Red]($#,##0.00)")).toBe("C2-")
    expect(formatCode("0%")).toBe("P0")
    expect(formatCode("0.00E+00")).toBe("S2")
    expect(formatCode("m/d/yyyy")).toBe("D4")
    expect(formatCode("d-mmm-yy")).toBe("D1")
    expect(formatCode("h:mm AM/PM")).toBe("D7")
    expect(formatCode("h:mm:ss")).toBe("D8")
    expect(formatCode("(0.00)")).toBe("F2()")
  })

  it("reads the cell's format, alignment, protection and column width", () => {
    const range = { r0: 0, c0: 1, r1: 0, c1: 1 }
    const fmt: Op = { kind: "format", sheet: 0, ranges: [range], patch: { numberFormat: "0.00%", locked: false } }
    expect(run('=CELL("format",B1)', [fmt])).toBe("P2")
    expect(run('=CELL("protect",B1)', [fmt])).toBe(0)
    const right: Op = { kind: "format", sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], patch: { hAlign: "right" } }
    expect(run('=CELL("prefix",A1)', [right])).toBe('"')
    expect(run('=CELL("width",A1)')).toEqual([[8, true]])
    const wide: Op = { kind: "axis", sheet: 0, axis: "cols", entries: [{ index: 0, size: 145, hidden: false }] }
    expect(run('=CELL("width",A1)', [wide])).toEqual([[20, false]])
  })

  it("gives the address, with the sheet when it's another one", () => {
    expect(run('=CELL("address",B2:B3)')).toBe("$B$2")
    expect(run('=CELL("address",Notes!C4)')).toBe("Notes!$C$4")
    expect(run('=CELL("contents",B2)')).toBe(-3.5)
    expect(run('=CELL("type",A1)')).toBe("l")
  })
})

describe("INFO", () => {
  it("counts the sheets and knows the recalculation mode", () => {
    expect(run('=INFO("numfile")')).toBe(2)
    expect(run('=INFO("recalc")')).toBe("Automatic")
    expect(run('=INFO("bogus")')).toEqual(err("#VALUE!"))
  })
})

describe("FILTERXML", () => {
  const xml =
    "<shop><item kind='pen'><name>Gel pen</name><price>2.5</price></item><item kind='ink'><name>Blue ink</name><price>4</price></item></shop>"
  it("finds elements, attributes and text by XPath", () => {
    expect(run(`=FILTERXML("${xml}","//name")`)).toEqual([["Gel pen"], ["Blue ink"]])
    expect(run(`=FILTERXML("${xml}","//item[@kind='ink']/price")`)).toBe(4)
    expect(run(`=FILTERXML("${xml}","//item[price>3]/name")`)).toBe("Blue ink")
    expect(run(`=FILTERXML("${xml}","//item[contains(name,'pen')]/@kind")`)).toBe("pen")
    expect(run(`=FILTERXML("${xml}","/shop/item[last()]/name/text()")`)).toBe("Blue ink")
    expect(run(`=FILTERXML("${xml}","//price | //name")`)).toEqual([
      ["Gel pen"],
      ["2.5"].map(Number),
      ["Blue ink"],
      [4],
    ])
  })

  it("rejects broken XML and expressions that aren't nodes", () => {
    expect(run('=FILTERXML("<a><b></a>","//b")')).toEqual(err("#VALUE!"))
    expect(run('=FILTERXML("<a>&bogus;</a>","/a")')).toEqual(err("#VALUE!"))
    expect(run('=FILTERXML("<a><b>1</b></a>","count(//b)")')).toEqual(err("#VALUE!"))
  })
})
