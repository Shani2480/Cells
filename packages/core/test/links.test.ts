import { describe, expect, it } from "vitest"
import { linkKey, Workbook, type ExternalBook } from "../src"

const prices = (): ExternalBook => ({
  name: "Price List.xlsx",
  target: "Price List.xlsx",
  sheets: [
    {
      name: "Prices",
      cells: new Map<number, string | number>([
        [linkKey(0, 0), "Pens"],
        [linkKey(0, 1), 1.5],
        [linkKey(1, 0), "Ink"],
        [linkKey(1, 1), 4],
        [linkKey(2, 0), "Paper"],
        [linkKey(2, 1), 7.25],
      ]),
    },
  ],
  names: [{ name: "Rate", refersTo: "='Prices'!$B$2" }],
})

function book() {
  const wb = new Workbook()
  wb.transact([{ kind: "externalLinks", links: [prices()] }], { label: "Links" })
  const s = wb.sheets[0]
  const put = (row: number, input: string) =>
    wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row, col: 0, data: { input, value: 0 } }] }], {
      label: "Type",
    })
  return { wb, s, put }
}

describe("links to other workbooks", () => {
  it("calculate from the values last read from the linked file", () => {
    const { s, put } = book()
    put(0, "='[Price List.xlsx]Prices'!$B$1*2")
    put(1, "=SUM('[Price List.xlsx]Prices'!$B$1:$B$3)")
    put(2, "='Price List.xlsx'!Rate+1")
    put(3, "=VLOOKUP(\"Ink\",'[Price List.xlsx]Prices'!$A:$B,2,FALSE)")
    put(4, "=COUNTA('[Price List.xlsx]Prices'!A:A)")
    expect([0, 1, 2, 3, 4].map((r) => s.cells.value(r, 0))).toEqual([3, 12.75, 5, 4, 3])
  })

  it("unquoted when the names allow it; an unknown workbook or sheet is #REF!", () => {
    const { wb, s, put } = book()
    wb.transact([{ kind: "externalLinks", links: [{ ...prices(), name: "Prices.xlsx" }] }], { label: "Rename" })
    put(0, "=[Prices.xlsx]Prices!B2+1")
    put(1, "=[Other.xlsx]Prices!B2")
    put(2, "='[Prices.xlsx]Nope'!B2")
    expect([0, 1, 2].map((r) => s.cells.value(r, 0))).toEqual([5, { error: "#REF!" }, { error: "#REF!" }])
  })

  it("new values from the file recalculate the formulas; undo brings the old ones back", () => {
    const { wb, s, put } = book()
    put(0, "='[Price List.xlsx]Prices'!B2*10")
    const next = prices()
    next.sheets[0].cells.set(linkKey(1, 1), 4.5)
    wb.transact([{ kind: "externalLinks", links: [next] }], { label: "Update Values" })
    expect(s.cells.value(0, 0)).toBe(45)
    wb.undo()
    expect(s.cells.value(0, 0)).toBe(40)
  })
})
