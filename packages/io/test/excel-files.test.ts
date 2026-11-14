// Files Excel itself wrote (the made-up samples, opened and saved again by Excel 365): Cells reads them as it wrote
// them, and saving keeps the parts it doesn't read.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { passwordMatches, type CellStyle, type Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "../src/zip"
import { samples } from "./samples"

const fixture = (name: string) =>
  new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}.xlsx`, import.meta.url)))

/** Excel writes number formats in its own spelling: "-" and " " escaped. Both show the same. */
const canonical = (style: CellStyle) =>
  style.numberFormat ? { ...style, numberFormat: style.numberFormat.replace(/\\([- ])/g, "$1") } : style

/** Every stored cell: value, formula and format, by sheet. */
function dump(wb: Workbook) {
  return wb.sheets.map((s) => {
    const cells: unknown[] = []
    const used = s.cells.usedBounds()
    if (used)
      s.cells.forEach(used, (row, col, d) =>
        cells.push([row, col, d.value, d.spill ? undefined : d.input, canonical(wb.styles.get(s.styleAt(row, col)))]),
      )
    return { name: s.name, hidden: s.hidden, cells }
  })
}

describe.each(["budget", "formats", "conditional", "layout", "protected"])("Excel's %s.xlsx", (name) => {
  it("reads as the workbook Cells saved", async () => {
    const read = await openXlsx(fixture(name), `${name}.xlsx`)
    const original = samples[name]()
    expect(dump(read.workbook)).toEqual(dump(original))
    read.workbook.sheets.forEach((s, i) => {
      const o = original.sheets[i]
      expect(s.merges.all).toEqual(o.merges.all)
      expect(s.conditional).toEqual(o.conditional)
      expect(s.annotations.all()).toEqual(o.annotations.all())
      expect(s.protection).toEqual(o.protection)
      expect(s.tabColor).toEqual(o.tabColor)
      for (let c = 0; c < 20; c++) expect([c, s.cols.size(c)]).toEqual([c, o.cols.size(c)])
      // Rows Excel fitted to their fonts (no customHeight) are left out: Cells fits them in the app.
      for (let r = 0; r < 20; r++) if (!s.rows.isAuto(r)) expect([r, s.rows.size(r)]).toEqual([r, o.rows.size(r)])
    })
    expect(read.workbook.names.all()).toEqual(original.names.all())
    expect(read.workbook.structure.protection).toEqual(original.structure.protection)
  })

  it("saves back with Excel's own parts kept", async () => {
    const bytes = fixture(name)
    const read = await openXlsx(bytes, `${name}.xlsx`)
    const saved = await saveXlsx(read.workbook, {
      source: read.source,
      views: read.views,
      activeSheet: read.activeSheet,
    })
    const before = new Map((await readZip(bytes)).map((e) => [e.normalizedName, e.data]))
    const after = new Map((await readZip(saved)).map((e) => [e.normalizedName, e.data]))
    for (const kept of ["xl/theme/theme1.xml", "docProps/core.xml", "docProps/app.xml"])
      expect(after.get(kept), kept).toEqual(before.get(kept))
    expect(after.has("xl/calcChain.xml")).toBe(false)
    const again = await openXlsx(saved, `${name}.xlsx`)
    expect(dump(again.workbook)).toEqual(dump(read.workbook))
  })
})

describe("passwords Excel 365 saved (SHA-512, 100,000 spins)", () => {
  it("unlock with the right password only, and stay as they were when saved", async () => {
    const read = await openXlsx(fixture("passwords"), "passwords.xlsx")
    const sheet = read.workbook.sheets[0].protection!
    const book = read.workbook.structure.protection!
    expect(sheet.hash?.spinCount).toBe(100_000)
    const t = performance.now()
    expect(passwordMatches(sheet, "Secret 1!")).toBe(true)
    const ms = performance.now() - t
    expect(passwordMatches(sheet, "secret 1!")).toBe(false)
    expect(passwordMatches(book, "Book pw")).toBe(true)
    expect(passwordMatches(book, "Secret 1!")).toBe(false)
    expect(ms).toBeLessThan(1500)
    const saved = await openXlsx(await saveXlsx(read.workbook, { source: read.source }), "passwords.xlsx")
    expect(saved.workbook.sheets[0].protection).toEqual(sheet)
    expect(saved.workbook.structure.protection).toEqual(book)
  })
})
