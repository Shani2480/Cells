// Sorting checked against Excel 365 (Windows, en-US): fixtures/excel-sort.json is the order Excel gave the same
// values (ascending, descending, ascending with Case sensitive); the multi-level case below is Excel's own result.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { compareSortText, errorValue, SortError, sortOps, Workbook, type CellData, type CellValue } from "../src"

const excel = JSON.parse(readFileSync(new URL("./fixtures/excel-sort.json", import.meta.url), "utf8")) as Record<
  string,
  string[]
>

/** The probe's values as cells: numbers, booleans and errors as such, the rest text. */
function asValue(text: string): CellValue | undefined {
  if (text === "") return undefined
  if (["5", "-2.5", "100"].includes(text)) return Number(text)
  if (text === "TRUE" || text === "FALSE") return text === "TRUE"
  if (text === "#DIV/0!" || text === "#N/A") return errorValue(text)
  return text
}

const shown = (v: CellValue | undefined) =>
  v === undefined ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : typeof v === "object" ? v.error : String(v)

function sortColumn(values: string[], descending: boolean, matchCase = false) {
  const wb = new Workbook()
  const s = wb.sheets[0]
  values.forEach((v, r) => {
    const value = asValue(v)
    if (value !== undefined) s.cells.set(r, 0, { value })
    s.cells.set(r, 1, { value: r })
  })
  const range = { r0: 0, c0: 0, r1: values.length - 1, c1: 1 }
  wb.transact(sortOps(0, s, wb.styles, range, { keys: [{ index: 0, descending }], matchCase }), { label: "Sort" })
  return values.map((_, r) => shown(s.cells.get(r, 0)?.value))
}

describe("sort order, as Excel's", () => {
  it("ascending", () => expect(sortColumn(excel.values, false)).toEqual(excel.asc))
  it("descending (blanks still last, ties in their order)", () =>
    expect(sortColumn(excel.values, true)).toEqual(excel.desc))
  it("case sensitive (lowercase first)", () => expect(sortColumn(excel.values, false, true)).toEqual(excel.case))

  it("symbols before digits before letters, in Windows' order", () => {
    const order = [
      " x",
      "!x",
      '"x',
      "#x",
      "$x",
      "%x",
      "&x",
      "(x",
      ")x",
      "*x",
      ",x",
      ".x",
      "/x",
      ":x",
      ";x",
      "?x",
      "@x",
      "[x",
      "\\x",
      "]x",
      "^x",
      "_x",
      "`x",
      "{x",
      "|x",
      "}x",
      "~x",
      "€x",
      "+x",
      "<x",
      "=x",
      ">x",
      "§x",
      "°x",
      "0x",
      "½x",
      "1x",
      "²x",
      "ax",
      "'x",
      "-x",
    ]
    const shuffled = [...order].reverse()
    expect([...shuffled].sort((a, b) => compareSortText(a, b) || order.indexOf(a) - order.indexOf(b))).toEqual(order)
  })
})

describe("sorting a range", () => {
  /** Excel's result: Team ascending, then Score descending, with headers; row 5 hidden. */
  it("moves whole rows, shifts relative references, leaves hidden rows and outside formulas alone", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    const rows: (string | number)[][] = [
      ["Name", "Team", "Score", "Double", "Next", "Abs"],
      ["Ann", "Blue", 5, "=C2*2", "=C3", "=$C$2"],
      ["Bob", "Red", 7, "=C3*2", "=C4", "=$C$2"],
      ["Cy", "Blue", 9, "=C4*2", "=C5", "=$C$2"],
      ["Di", "red", 5, "=C5*2", "=C6", "=$C$2"],
      ["Ed", "Green", 1, "=C6*2", "=C7", "=$C$2"],
      ["Flo", "Blue", 3, "=C7*2", "=C8", "=$C$2"],
    ]
    rows.forEach((row, r) =>
      row.forEach((v, c) =>
        s.cells.set(r, c, typeof v === "string" && v.startsWith("=") ? { value: 0, input: v } : { value: v }),
      ),
    )
    s.cells.set(8, 0, { value: 0, input: "=C3" })
    s.rows.setHidden(4, true)
    const ops = sortOps(
      0,
      s,
      wb.styles,
      { r0: 0, c0: 0, r1: 6, c1: 5 },
      {
        header: true,
        keys: [{ index: 1 }, { index: 2, descending: true }],
      },
    )
    wb.transact(ops, { label: "Sort" })
    const text = (r: number) =>
      Array.from({ length: 6 }, (_, c) => {
        const d = s.cells.get(r, c)
        return d?.input ?? String(d?.value ?? "")
      }).join("|")
    expect([0, 1, 2, 3, 4, 5, 6, 8].map(text)).toEqual([
      "Name|Team|Score|Double|Next|Abs",
      "Cy|Blue|9|=C2*2|=C3|=$C$2",
      "Ann|Blue|5|=C3*2|=C4|=$C$2",
      "Flo|Blue|3|=C4*2|=C5|=$C$2",
      "Di|red|5|=C5*2|=C6|=$C$2",
      "Ed|Green|1|=C6*2|=C7|=$C$2",
      "Bob|Red|7|=C7*2|=C8|=$C$2",
      "=C3|||||",
    ])
    wb.undo()
    expect(text(1)).toBe("Ann|Blue|5|=C2*2|=C3|=$C$2")
  })

  it("left to right, by colour, by a custom list; notes go along; merged cells stop it", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    const red = wb.styles.intern({ fill: "#FF0000" })
    const put = (r: number, c: number, d: CellData) => s.cells.set(r, c, d)
    ;["Mar", "Jan", "Feb"].forEach((m, c) => put(0, c, { value: m }))
    ;[3, 1, 2].forEach((v, c) => put(1, c, { value: v, style: c === 2 ? red : undefined }))
    s.annotations.set(1, 0, { kind: "note", author: "Alex", text: "march", shown: false })
    const range = { r0: 0, c0: 0, r1: 1, c1: 2 }
    wb.transact(
      sortOps(0, s, wb.styles, range, { byColumns: true, keys: [{ index: 0, list: ["Jan", "Feb", "Mar"] }] }),
      {
        label: "Sort",
      },
    )
    expect([0, 1, 2].map((c) => s.cells.get(0, c)?.value)).toEqual(["Jan", "Feb", "Mar"])
    expect(s.annotations.get(1, 2)).toMatchObject({ text: "march" })
    expect(s.annotations.get(1, 0)).toBeUndefined()
    // The red cell first.
    wb.transact(
      sortOps(0, s, wb.styles, range, { byColumns: true, keys: [{ index: 1, on: "cellColor", color: "#ff0000" }] }),
      {
        label: "Sort",
      },
    )
    expect(s.cells.get(1, 0)).toMatchObject({ value: 2, style: red })
    s.merges.add({ r0: 0, c0: 0, r1: 0, c1: 1 })
    expect(() => sortOps(0, s, wb.styles, range, { keys: [{ index: 0 }] })).toThrow(SortError)
  })

  it("already in order: nothing to do", () => {
    const wb = new Workbook()
    ;[1, 2, 3].forEach((v, r) => wb.sheets[0].cells.set(r, 0, { value: v }))
    expect(sortOps(0, wb.sheets[0], wb.styles, { r0: 0, c0: 0, r1: 2, c1: 0 }, { keys: [{ index: 0 }] })).toEqual([])
  })
})
