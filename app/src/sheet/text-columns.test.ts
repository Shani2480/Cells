import { describe, expect, it, vi } from "vitest"
import { EN_US, regionalFrom, setRegional } from "@cells/core"
import { SheetController } from "./controller"
import { TEXT_COLUMNS_DEFAULT } from "./text-columns"

const comma = { ...TEXT_COLUMNS_DEFAULT, delimiters: [","] }

function type(c: SheetController, row: number, col: number, text: string) {
  c.goTo({ r0: row, c0: col, r1: row, c1: col })
  c.beginEdit("enter", text)
  c.commitEdit()
}
function withColumn(values: string[]) {
  const c = new SheetController()
  c.shell.alert = vi.fn()
  values.forEach((v, r) => {
    if (v !== "") type(c, r, 0, v)
  })
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  return c
}
const shown = (c: SheetController, rows: number, cols: number) =>
  Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, col) => {
      const d = c.sheet.cells.get(r, col)
      return d?.input ?? (d?.value === undefined ? "" : String(d.value))
    }),
  )

describe("Text to Columns (as Excel does it)", () => {
  it("clears as many columns as the widest row fills, on every row", () => {
    const c = withColumn(["a,,b", "x", ""])
    for (let r = 0; r < 3; r++) for (let col = 1; col < 4; col++) type(c, r, col, "old")
    c.data.textToColumns({ r0: 0, c0: 0, r1: 2, c1: 0 }, comma, { row: 0, col: 0 })
    expect(shown(c, 3, 4)).toEqual([
      ["a", "", "b", "old"],
      ["x", "", "", "old"],
      ["", "", "", "old"],
    ])
  })

  it("splits a formula's text, and reads each field as if typed", () => {
    const c = withColumn(["1,2026-03-04,007, 5 ,TRUE,'x,=1+2"])
    c.data.textToColumns({ r0: 0, c0: 0, r1: 0, c1: 0 }, comma, { row: 0, col: 0 })
    const row = [0, 1, 2, 3, 4, 5, 6].map((col) => c.sheet.cells.get(0, col))
    expect(row.map((d) => d?.value)).toEqual([1, 46085, 7, 5, true, "'x", 3])
    expect(c.wb.styles.get(row[1]?.style ?? 0).numberFormat).toBe("m/d/yyyy")
    expect(row[6]?.input).toBe("=1+2")
  })

  it("quotes and consecutive delimiters; fixed width trims the fields", () => {
    const c = withColumn(['"a,b",c', "d,,e"])
    c.data.textToColumns({ r0: 0, c0: 0, r1: 1, c1: 0 }, { ...comma, consecutive: true }, { row: 0, col: 0 })
    expect(shown(c, 2, 3)).toEqual([
      ["a,b", "c", ""],
      ["d", "e", ""],
    ])
    const f = withColumn(["abc  def", "ab   d   x"])
    f.data.textToColumns(
      { r0: 0, c0: 0, r1: 1, c1: 0 },
      { ...TEXT_COLUMNS_DEFAULT, kind: "fixed", breaks: [5] },
      {
        row: 0,
        col: 0,
      },
    )
    expect(shown(f, 2, 2)).toEqual([
      ["abc", "def"],
      ["ab", "d   x"],
    ])
  })

  it("a cell that doesn't split keeps its format; another destination leaves the source", () => {
    const c = withColumn(["1234.5", "a b  c"])
    c.data.textToColumns({ r0: 0, c0: 0, r1: 0, c1: 0 }, comma, { row: 0, col: 0 })
    expect(c.sheet.cells.get(0, 0)?.value).toBe(1234.5)
    c.data.textToColumns(
      { r0: 1, c0: 0, r1: 1, c1: 0 },
      { ...TEXT_COLUMNS_DEFAULT, delimiters: [" "] },
      {
        row: 2,
        col: 2,
      },
    )
    expect(c.sheet.cells.get(1, 0)?.value).toBe("a b  c")
    expect(shown(c, 3, 6)[2]).toEqual(["", "", "a", "b", "", "c"])
  })

  it("skipped columns, Text and asking before replacing data; one undo step", () => {
    const c = withColumn(["7,8,9"])
    type(c, 0, 3, "keep?")
    c.shell.confirm = vi.fn()
    const o = { ...comma, columns: ["text", "skip"] as const }
    c.data.textToColumns({ r0: 0, c0: 0, r1: 0, c1: 0 }, { ...o, columns: [...o.columns] }, { row: 0, col: 2 })
    expect(c.shell.confirm).toHaveBeenCalledWith(
      "Do you want to replace the contents of the destination cells?",
      expect.any(Function),
    )
    ;(c.shell.confirm as ReturnType<typeof vi.fn>).mock.calls[0][1]()
    expect(c.sheet.cells.get(0, 2)?.value).toBe("7")
    expect(c.sheet.cells.get(0, 3)?.value).toBe(9)
    c.undo()
    expect(c.sheet.cells.get(0, 3)?.value).toBe("keep?")
  })

  it("only one column at a time, and not an empty one", () => {
    const c = withColumn(["a,b"])
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 1 })
    c.data.openTextToColumns()
    expect(c.shell.alert).toHaveBeenCalledWith(expect.stringMatching(/only one column at a time/))
    const e = withColumn([])
    e.data.openTextToColumns()
    expect(e.shell.alert).toHaveBeenCalledWith("No data was selected to parse.")
    c.shell.textToColumns = vi.fn()
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.data.openTextToColumns()
    expect(c.shell.textToColumns).toHaveBeenCalledWith({ r0: 0, c0: 0, r1: 0, c1: 0 }, ["a,b"])
  })
})

describe("Text to Columns in a German regional setting", () => {
  it("reads the fields with the region's separators and date order", () => {
    setRegional(regionalFrom({ dateFormat: "dd.MM.yyyy", decimalSeparator: ",", groupSeparator: "." }))
    try {
      const c = withColumn(["1,5;31.01.2026;1.234"])
      const o = c.data.textColumns
      expect(o).toMatchObject({ decimal: ",", thousands: "." })
      c.data.textToColumns({ r0: 0, c0: 0, r1: 0, c1: 0 }, { ...o, delimiters: [";"] }, { row: 0, col: 0 })
      expect([0, 1, 2].map((col) => c.sheet.cells.value(0, col))).toEqual([1.5, 46053, 1234])
    } finally {
      setRegional(EN_US)
    }
  })
})
