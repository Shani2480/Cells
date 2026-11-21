// macros.xlsm / macros.xls: a workbook with a standard module ("Tools") and code in ThisWorkbook, as Excel 365 saved
// it (made in the Office VM). Cells lists the macros and shows the code; it doesn't run them.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { openXls, openXlsx, saveXlsx } from "../src"
import { decompressVba, vbaMacros } from "../src/vba"

const file = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}`, import.meta.url)))

describe("VBA projects", () => {
  it("decompresses MS-OVBA's example (2.4.1.3.2)", () => {
    const hex =
      "01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 30 6D 6E 6F 70 06 71 02 70 04 10 72 73 74 75 76 10 77 78 79 7A 00 3C"
    const bytes = Uint8Array.from(hex.split(" ").map((h) => parseInt(h, 16)))
    expect(new TextDecoder().decode(decompressVba(bytes))).toBe(
      "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
    )
  })

  it("reads the modules of an .xlsm and lists the macros as Excel does", async () => {
    const { workbook } = await openXlsx(file("macros.xlsm"), "macros.xlsm")
    const byName = new Map(workbook.vba.map((m) => [m.name, m]))
    expect(byName.get("Tools")?.kind).toBe("module")
    expect(byName.get("ThisWorkbook")?.kind).toBe("document")
    expect(byName.get("Tools")?.code).toContain('Range("A1:A5").Value = Date')
    expect(byName.get("Tools")?.code).not.toMatch(/Attribute VB_/)
    expect(
      vbaMacros(workbook.vba)
        .map((m) => m.name)
        .sort(),
    ).toEqual(["ClearSheet", "FillDates", "ThisWorkbook.Greet"])
  })

  it("reads them from an .xls too, and an .xlsm keeps its project when saved", async () => {
    const { workbook } = await openXls(file("macros.xls"), "macros.xls")
    expect(
      vbaMacros(workbook.vba)
        .map((m) => m.name)
        .sort(),
    ).toEqual(["ClearSheet", "FillDates", "ThisWorkbook.Greet"])
    const read = await openXlsx(file("macros.xlsm"), "macros.xlsm")
    const saved = await saveXlsx(read.workbook, { source: read.source, fileName: "macros.xlsm" })
    expect((await openXlsx(saved, "macros.xlsm")).workbook.vba).toHaveLength(read.workbook.vba.length)
  })
})
