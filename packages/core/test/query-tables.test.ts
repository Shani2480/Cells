import { describe, expect, it } from "vitest"
import { queryTableName, Workbook, type QueryTable } from "../src"

const link = (name: string, r0 = 0): QueryTable => ({
  name,
  connection: name,
  description: "",
  range: { r0, c0: 0, r1: r0 + 3, c1: 1 },
  source: "/home/alex/data.csv",
  options: {
    kind: "delimited",
    delimiters: [","],
    qualifier: '"',
    consecutive: false,
    breaks: [],
    startRow: 1,
    columns: [],
    decimal: ".",
    thousands: ",",
  },
  refreshOnLoad: false,
})

describe("query tables", () => {
  it("names from a file, valid and unique", () => {
    expect(queryTableName("Sales 2026.csv", [])).toBe("Sales_2026")
    expect(queryTableName("2026.csv", [])).toBe("_2026")
    expect(queryTableName("AB12.txt", [])).toBe("_AB12")
    expect(queryTableName("sales.csv", ["Sales", "sales_1"])).toBe("sales_2")
  })

  it("a change is one step to undo; rows inserted above move the range, deleting all of it drops the link", () => {
    const wb = new Workbook()
    wb.transact([{ kind: "queryTables", sheet: 0, tables: [link("a"), link("b", 10)] }], { label: "Link" })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 5, count: 2 }], { label: "Insert" })
    expect(wb.sheets[0].queryTables.map((q) => q.range.r0)).toEqual([0, 12])
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 12, count: 4 }], { label: "Delete" })
    expect(wb.sheets[0].queryTables.map((q) => q.name)).toEqual(["a"])
    wb.undo()
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].queryTables).toEqual([])
  })
})
