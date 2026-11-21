// shared.xlsx: a legacy shared workbook (Share Workbook) as Excel 365 saved it (made in the Office VM), with a
// change history of one edit.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/shared.xlsx", import.meta.url)))
const parts = async (data: Uint8Array) => (await readZip(data)).map((e) => e.name)

describe("legacy shared workbooks", () => {
  it("knows the file is shared and keeps its change history while it stays shared", async () => {
    const { workbook, source } = await openXlsx(bytes, "shared.xlsx")
    expect(workbook.legacyShared).toBe(true)
    const saved = await saveXlsx(workbook, { source })
    expect((await parts(saved)).filter((n) => n.includes("revisions/"))).toHaveLength(5)
    expect((await openXlsx(saved, "shared.xlsx")).workbook.legacyShared).toBe(true)
  })

  it("Unshare Workbook drops the history and user names when saving; the cells stay", async () => {
    const { workbook, source } = await openXlsx(bytes, "shared.xlsx")
    workbook.legacyShared = false
    const saved = await saveXlsx(workbook, { source })
    expect((await parts(saved)).some((n) => n.includes("revisions/"))).toBe(false)
    const zip = await readZip(saved)
    const rels = new TextDecoder().decode(zip.find((e) => e.name === "xl/_rels/workbook.xml.rels")!.data)
    expect(rels).not.toMatch(/revisionHeaders|usernames/)
    const again = await openXlsx(saved, "shared.xlsx")
    expect(again.workbook.legacyShared).toBe(false)
    expect(again.workbook.sheets[0].cells.value(0, 1)).toBe(15)
  })
})
