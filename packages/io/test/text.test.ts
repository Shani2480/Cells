// CSV and text files: reading as Excel's Open does (checked against Excel 365 opening fixtures/excel/text/probe.csv)
// and saving byte for byte as Excel's Save As writes (the excel-* fixtures, saved by Excel from the same cells).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { EN_US, errorValue, regionalFrom, setRegional, Workbook, type CellData } from "@cells/core"
import {
  dateIn,
  decodeText,
  detectBreaks,
  detectDelimiter,
  encodeText,
  openText,
  parseDelimited,
  saveText,
  sheetNameFor,
  sniffEncoding,
  splitText,
  defaultImport,
  CSV,
} from "../src"

const fixture = (name: string) =>
  new Uint8Array(readFileSync(new URL(`./fixtures/excel/text/${name}`, import.meta.url)))

const cellAt = (wb: Workbook, ref: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1
  const data = wb.sheets[0].cells.get(Number(m[2]) - 1, col)
  if (!data) return undefined
  const nf = data.style ? wb.styles.get(data.style).numberFormat : undefined
  return { value: data.value, input: data.input, nf }
}

describe("splitting", () => {
  it("handles quotes, doubled quotes, line breaks in fields and every line end", () => {
    expect(parseDelimited('a,"b,c","say ""hi"""\r\n"x\r\ny",z\rlast\n', CSV)).toEqual([
      ["a", "b,c", 'say "hi"'],
      ["x\ny", "z"],
      ["last"],
    ])
    // Text after a closing quote joins the field; a quote inside an unquoted field is just a character.
    expect(parseDelimited('"ab"cd,e"f', CSV)).toEqual([["abcd", 'e"f']])
    // Empty lines are empty rows; a last line break isn't a row; an unclosed quote runs to the end.
    expect(parseDelimited("a\n\nb\n", CSV)).toEqual([["a"], [""], ["b"]])
    expect(parseDelimited('a,"open\nstill', CSV)).toEqual([["a", "open\nstill"]])
    expect(parseDelimited("a,,b,", CSV)).toEqual([["a", "", "b", ""]])
  })

  it("treats consecutive delimiters as one when asked, and takes several delimiters", () => {
    const o = { delimiters: [" ", ";"], qualifier: '"', consecutive: true }
    expect(parseDelimited("a   b;;c", o)).toEqual([["a", "b", "c"]])
    expect(parseDelimited("a;'b;c'", { delimiters: [";"], qualifier: "'", consecutive: false })).toEqual([["a", "b;c"]])
    expect(parseDelimited('a;"b', { delimiters: [";"], qualifier: null, consecutive: false })).toEqual([["a", '"b']])
  })

  it("finds the delimiter a file uses", () => {
    expect(detectDelimiter("name;amount\nRent;1200,50\nFood;310\n")).toBe(";")
    expect(detectDelimiter("a,b,c\n1,2,3\n")).toBe(",")
    expect(detectDelimiter("a\tb\n1\t2\n")).toBe("\t")
    expect(detectDelimiter('"x, y";1\n"z, w";2\n')).toBe(";")
    expect(detectDelimiter("just one column\nand another\n")).toBe(",")
  })

  it("fixed width: finds the columns and cuts lines there", () => {
    const lines = ["Name      Qty  Price", "Apples      3   1.20", "Pears      12   0.85"]
    expect(detectBreaks(lines)).toEqual([10, 15])
    const o = { ...defaultImport("", "x.txt"), kind: "fixed" as const, breaks: [10, 15] }
    expect(splitText(lines.join("\n"), o)).toEqual([
      ["Name", "Qty", "Price"],
      ["Apples", "3", "1.20"],
      ["Pears", "12", "0.85"],
    ])
  })
})

describe("encodings", () => {
  it("reads the byte-order mark, and tells UTF-8 from Windows-1252 without one", () => {
    expect(sniffEncoding(new Uint8Array([0xef, 0xbb, 0xbf, 0x61]))).toEqual({ encoding: "utf-8", bom: true })
    expect(sniffEncoding(new Uint8Array([0xff, 0xfe, 0x61, 0]))).toEqual({ encoding: "utf-16le", bom: true })
    expect(sniffEncoding(new Uint8Array([0x61, 0, 0x62, 0]))).toEqual({ encoding: "utf-16le", bom: false })
    expect(sniffEncoding(new TextEncoder().encode("café"))).toEqual({ encoding: "utf-8", bom: false })
    expect(sniffEncoding(new Uint8Array([0x63, 0x61, 0x66, 0xe9]))).toEqual({ encoding: "windows-1252", bom: false })
    expect(decodeText(new Uint8Array([0x80, 0x20, 0xe9, 0x93, 0x94])).text).toBe("€ é“”")
  })

  it("writes each encoding, with ? for what Windows-1252 lacks (one per UTF-16 unit, as Excel)", () => {
    for (const encoding of ["utf-8", "utf-16le", "utf-16be", "windows-1252"] as const) {
      const text = "Café € “quoted”"
      expect(decodeText(encodeText(text, encoding, true), encoding).text).toBe(text)
    }
    expect([...encodeText("中😀é", "windows-1252", false)]).toEqual([0x3f, 0x3f, 0x3f, 0xe9])
  })
})

describe("opening a CSV as Excel does", () => {
  it("parses each field as if typed (Excel 365, en-US)", async () => {
    const { workbook: wb, warnings } = await openText(fixture("probe.csv"), "probe.csv")
    expect(warnings).toEqual([])
    expect(wb.sheets[0].name).toBe("probe")
    // What Excel showed for each cell (Formula, Value2 and NumberFormat).
    const expected: Record<string, Partial<{ value: unknown; input: string; nf: string }>> = {
      A2: { value: "abc" },
      B2: { value: 1234.5 },
      C2: { value: 7 },
      D2: { value: 46095, nf: "m/d/yyyy" },
      E2: { value: 1, nf: "$#,##0_);[Red]($#,##0)" },
      F2: { value: 234.5 },
      G2: { value: 0.12, nf: "0%" },
      H2: { value: 'a,"b"' },
      I2: { value: true },
      J2: { value: 3, input: "=1+2" },
      K2: { value: 100000, nf: "0.00E+00" },
      L2: { value: 123456789012345000 },
      M2: { value: -42 },
      N2: { value: 1.5, nf: "# ?/?" },
      O2: { nf: "h:mm" },
      P2: { value: "31/12/2026" },
      Q2: { value: "2026-03-14T10:00" },
      A3: { value: "line1\nline2" },
      B3: { value: 0 },
      C3: { value: 12 },
      D3: { value: 46095, nf: "m/d/yyyy" },
      E3: { value: 1234, nf: "#,##0" },
      F3: { value: 0.5 },
      H3: { value: false },
      I3: { value: 1, input: "=SUM(1)" },
      J3: { value: "1E+400" },
      K3: { value: 1234567890123 },
      L3: { value: "-" },
      M3: { value: -3, input: "=-3" },
      N3: { nf: "h:mm AM/PM" },
      O3: { value: 46048, nf: "d-mmm" },
      // Excel read the file as Windows-1252 and garbled this line; Cells sees it is UTF-8.
      A4: { value: "é ü 中文" },
      C4: { value: "x" },
    }
    for (const [ref, want] of Object.entries(expected)) expect(cellAt(wb, ref), ref).toMatchObject(want)
    expect(cellAt(wb, "O2")!.value).toBeCloseTo(0.572916666666667, 12)
    expect(cellAt(wb, "N3")!.value).toBeCloseTo(0.878472222222222, 12)
    // An empty quoted field is an empty cell.
    expect(cellAt(wb, "G3")).toBeUndefined()
    expect(wb.sheets[0].cells.usedBounds()).toEqual({ r0: 0, c0: 0, r1: 3, c1: 16 })
  })

  it("reads the wizard's column formats: text, skipped, dates in any order, other separators", async () => {
    const csv = "id;when;amount;note\n007;31.12.2026;1.234,50;x\n010;1.2.26 8:30;-5,5;y\n"
    const { workbook: wb } = await openText(new TextEncoder().encode(csv), "export.txt", {
      delimiters: [";"],
      columns: ["text", "DMY", "general", "skip"],
      decimal: ",",
      thousands: ".",
      startRow: 2,
    })
    expect(cellAt(wb, "A1")).toMatchObject({ value: "007", nf: "@" })
    expect(cellAt(wb, "B1")).toMatchObject({ value: 46387, nf: "m/d/yyyy" })
    expect(cellAt(wb, "B2")!.nf).toBe("m/d/yyyy h:mm")
    expect(cellAt(wb, "B2")!.value).toBeCloseTo(46054 + 8.5 / 24, 9)
    expect(cellAt(wb, "C1")!.value).toBe(1234.5)
    expect(cellAt(wb, "C2")!.value).toBe(-5.5)
    expect(cellAt(wb, "D1")).toBeUndefined()
  })

  it("dates in each order, with two-digit years and month names", () => {
    expect(dateIn("2026-03-14", "YMD")?.value).toBe(46095)
    expect(dateIn("14/03/2026", "DMY")?.value).toBe(46095)
    expect(dateIn("03.14.26", "MDY")?.value).toBe(46095)
    expect(dateIn("14 Mar 2026", "DMY")?.value).toBe(46095)
    expect(dateIn("1/1/30", "MDY")?.value).toBe(10959)
    expect(dateIn("31/02/2026", "DMY")).toBeNull()
    expect(dateIn("hello", "DMY")).toBeNull()
  })

  it("names the sheet after the file", () => {
    expect(sheetNameFor("Q1 [draft]: sales.csv")).toBe("Q1 _draft__ sales")
    expect(sheetNameFor("a-very-long-file-name-for-the-export-2026.csv")).toBe("a-very-long-file-name-for-the-e")
  })

  it("warns when a file has more rows than a sheet", async () => {
    const big = new TextEncoder().encode("1\n".repeat(1_048_577))
    const { workbook: wb, warnings } = await openText(big, "big.csv", { calculate: false })
    expect(wb.sheets[0].cells.get(1_048_575, 0)?.value).toBe(1)
    expect(warnings[0]).toMatch(/1,048,577 rows/)
  }, 30_000)
})

describe("saving as Excel does", () => {
  /** The cells Excel saved in the excel-* fixtures. */
  function excelSheet() {
    const wb = new Workbook()
    const s = wb.sheets[0]
    const nf = (numberFormat: string) => wb.styles.intern({ numberFormat })
    const row: [number, CellData][] = [
      [0, { value: "plain" }],
      [1, { value: 'a,"b"' }],
      [2, { value: "line1\nline2" }],
      [3, { value: 1234567.891 }],
      [4, { value: 123456789012 }],
      [5, { value: 0.1 + 0.2 }],
      [6, { value: 46095.5, style: nf("m/d/yyyy h:mm") }],
      [7, { value: 0.125, style: nf("0.0%") }],
      [8, { value: -1234.5, style: nf("$#,##0.00_);[Red]($#,##0.00)") }],
      [9, { value: errorValue("#DIV/0!"), input: "=1/0" }],
      [10, { value: true }],
      [11, { value: "é ü 中文 😀" }],
      [12, { value: " lead space" }],
      [13, { value: "semi;colon" }],
      [14, { value: "tab\there" }],
      [15, { value: 12345678901234567000 }],
      [16, { value: "quote", input: "'quote" }],
      [17, { value: 5, style: nf('"Total: "0') }],
    ]
    for (const [col, data] of row) s.cells.set(0, col, data)
    s.cells.set(2, 1, { value: "row3" })
    return wb
  }

  it("CSV UTF-8 matches Excel's file byte for byte", () => {
    const ours = saveText(excelSheet(), 0, { delimiter: ",", encoding: "utf-8", bom: true })
    expect(new TextDecoder().decode(ours)).toBe(new TextDecoder().decode(fixture("excel-utf8.csv")))
    expect([...ours]).toEqual([...fixture("excel-utf8.csv")])
  })

  it("Unicode Text (UTF-16, tab-delimited) matches Excel's file byte for byte", () => {
    const ours = saveText(excelSheet(), 0, { delimiter: "\t", encoding: "utf-16le", bom: true })
    expect([...ours]).toEqual([...fixture("excel-unicode.txt")])
  })

  it("CSV and Text (Windows-1252) match Excel's files where the characters fit", () => {
    const wb = excelSheet()
    wb.sheets[0].cells.set(0, 11, { value: "é ü ?? ??" })
    expect([...saveText(wb, 0, { delimiter: ",", encoding: "windows-1252", bom: false })]).toEqual([
      ...fixture("excel-ansi.csv"),
    ])
    expect([...saveText(wb, 0, { delimiter: "\t", encoding: "windows-1252", bom: false })]).toEqual([
      ...fixture("excel-tab.txt"),
    ])
  })

  it("a Windows-1252 file that gets characters it can't hold is saved as UTF-8 instead of losing them", () => {
    const bytes = saveText(excelSheet(), 0, { delimiter: ",", encoding: "windows-1252", bom: false })
    expect(decodeText(bytes)).toMatchObject({ encoding: "utf-8", bom: true })
    expect(decodeText(bytes).text).toContain("中文 😀")
  })

  it("an empty sheet is an empty file, and saving what was opened gives the same text", async () => {
    expect(saveText(new Workbook(), 0, { delimiter: ",", encoding: "utf-8", bom: false }).length).toBe(0)
    const csv = 'name,amount,when\r\nRent,1200.5,3/14/2026\r\n"Food, drink",310,\r\n'
    const { workbook } = await openText(new TextEncoder().encode(csv), "budget.csv")
    expect(new TextDecoder().decode(saveText(workbook, 0, { delimiter: ",", encoding: "utf-8", bom: false }))).toBe(csv)
  })

  it("opens and saves the same with a German regional setting (CSV stays en-US)", async () => {
    setRegional(regionalFrom({ dateFormat: "dd.MM.yyyy", decimalSeparator: ",", groupSeparator: "." }))
    try {
      const csv = 'name,amount,when\r\nRent,1200.5,3/14/2026\r\n"Food, drink",310,\r\n'
      const { workbook } = await openText(new TextEncoder().encode(csv), "budget.csv")
      expect(workbook.sheets[0].cells.value(1, 1)).toBe(1200.5)
      expect(workbook.sheets[0].cells.value(1, 2)).toBe(46095)
      expect(new TextDecoder().decode(saveText(workbook, 0, { delimiter: ",", encoding: "utf-8", bom: false }))).toBe(
        csv,
      )
    } finally {
      setRegional(EN_US)
    }
  })
})
