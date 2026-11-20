// GETPIVOTDATA against the PivotTables Excel made: each formula and the answer Excel 365 gave it (formulas typed
// into Excel over the same fixtures, the results read back).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { errorValue, type CellValue } from "@cells/core"
import { openXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}`, import.meta.url)))
const REF = errorValue("#REF!")

const CASES: Record<string, [string, CellValue][]> = {
  "pivot-basic.xlsx": [
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3)', 6630],
    ['=GETPIVOTDATA("Amount",Compact!A3)', 6630],
    ['=GETPIVOTDATA("sum of amount",Compact!B9,"region","north")', 1672.5],
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3,"Region","North","Product","Pears")', 945],
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3,"Product","Pears")', REF],
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3,"Region","Nowhere")', REF],
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3,"Channel","Shop")', REF],
    ['=GETPIVOTDATA("Sum of Amount",Compact!Z99)', REF],
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3:B20)', 6630],
    ['=GETPIVOTDATA("Count of Units",Compact!A3)', REF],
    ['=GETPIVOTDATA("Count of Units",Outline!A3,"Region","East")', 6],
    ['=GETPIVOTDATA("Units",Outline!A3,"Region","East")', 6],
    ['=GETPIVOTDATA("Sum of Units",Cross!A3,"Channel","Online")', 68],
    ['=GETPIVOTDATA("Sum of Units",Cross!A3,"Channel","Shop")', REF],
    ['=GETPIVOTDATA("Sum of Units",Cross!A3,"Product","Apples","Region","East")', REF],
    ['=GETPIVOTDATA("Sum of Units",Cross!A3,"Region","South","Product","Apples")', 15],
    ['=GETPIVOTDATA("Sum of Units",Cross!A1)', REF],
    ['=GETPIVOTDATA("StdDev of Amount",Functions!A3,"Region","East")', 193.555009407317],
    ['=GETPIVOTDATA("Sum of Units",ValuesRows!A3,"Region","East","Product","Pears")', 23],
    ['=GETPIVOTDATA("Sum of Units",ValuesRows!A3,"Region","East")', 75],
    ["=GETPIVOTDATA(Compact!B3,Compact!A3,Compact!A4,Compact!A5)", REF],
    ['=GETPIVOTDATA("Sum of Amount",Compact!A3,"Region","North ")', REF],
    ['=GETPIVOTDATA("Average of Units",Tabular!A3,"Region","East","Product","Pears")', 11.5],
  ],
  "pivot-more.xlsx": [
    ['=GETPIVOTDATA("Sum of Units",TwoPages!A3,"Region","East")', REF],
    ['=GETPIVOTDATA("Sum of Units",TwoPages!A3,"Product","Pears")', 44],
    ['=GETPIVOTDATA("Sum of Amount",Numbers!A3,"Units",10)', 357.5],
    ['=GETPIVOTDATA("Sum of Amount",Numbers!A3,"Units","10")', 357.5],
    ['=GETPIVOTDATA("Sum of Units",Subtotals!A3,"Region","East","Product","Pears")', 34],
    ['=GETPIVOTDATA("Sum of Units",Subtotals!A3,"Region","East")', REF],
    ['=GETPIVOTDATA("Sum of Units",Subtotals!A3,"Region","East","Product","Pears","Channel","Shop")', 34],
    ['=GETPIVOTDATA("Average of Amount",Collapsed!A3,"Area","North")', 278.75],
    ['=GETPIVOTDATA("Average of Amount",Collapsed!A3,"Region","North")', REF],
    ['=GETPIVOTDATA("Average of Amount",Collapsed!A3,"Area","North","Product","Apples")', REF],
    ['=GETPIVOTDATA("Average of Amount",Collapsed!A3,"Area","East","Product","Apples")', 335],
    ['=GETPIVOTDATA("Average of Amount",Collapsed!A3,"Area","South","Product","Apples")', 147.5],
    ['=GETPIVOTDATA("Average of Amount",Collapsed!A3,"Area","West","Product","Apples")', 397.5],
    ['=GETPIVOTDATA("Count of Units2",ColOnly!A3,"Product","Pears")', 9],
    ['=GETPIVOTDATA("Max of Amount",TwoCols!A3,"Region","East","Product","Pears")', REF],
    ['=GETPIVOTDATA("Max of Amount",TwoCols!A3,"Product","Pears")', REF],
    ['=GETPIVOTDATA("Max of Amount",TwoCols!A3,"Product","Pears","Channel","Shop")', 597.5],
  ],
  "pivot-showas.xlsx": [
    ['=GETPIVOTDATA("Amount PctOf",PctOf!A3,"Region","East","Channel","Online")', errorValue("#NULL!")],
    ['=GETPIVOTDATA("Amount PctOf",PctOf!A3,"Region","North","Channel","Online")', 0],
    ['=GETPIVOTDATA("Amount PctOf",PctOf!A3,"Region","East","Channel","Shop")', 1.14947683109118],
    ['=GETPIVOTDATA("Amount PctParent",PctParent!A3)', 0],
    ['=ISBLANK(GETPIVOTDATA("Amount PctOf",PctOf!A3,"Region","North","Channel","Online"))', true],
    ['=GETPIVOTDATA("Amount PctTotal",PctTotal!A3,"Region","East","Product","Apples","Channel","Online")', 0],
  ],
}

describe("GETPIVOTDATA", () => {
  it.each(Object.keys(CASES))("gives Excel's answers over %s", async (name) => {
    const { workbook: wb } = await openXlsx(fixture(name), name)
    const s = wb.sheets.length
    wb.transact([{ kind: "addSheet", sheet: s, name: "F" }], { label: "Add", history: false })
    const list = CASES[name]
    wb.transact(
      [{ kind: "setCells", sheet: s, cells: list.map(([f], row) => ({ row, col: 0, data: { input: f, value: 0 } })) }],
      { label: "Set", history: false },
    )
    wb.calculate(true)
    list.forEach(([f, want], row) => {
      const got = wb.sheets[s].cells.value(row, 0) ?? null
      if (typeof want === "number") expect(got, f).toBeCloseTo(want, 9)
      else expect(got, f).toEqual(want)
    })
  })
})
