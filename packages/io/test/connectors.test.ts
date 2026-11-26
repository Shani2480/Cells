// A connector glued to two shapes, as Excel 365 saved it (fixtures/excel/connector.xlsx): an arrow from Rectangle 1's
// right side (site 4 in Excel's counting, idx 3) to Oval 2's left (site 3, idx 2).
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { isChart, type SheetObject, type SheetShape } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/connector.xlsx", import.meta.url)))
const shapes = (list: SheetObject[]) => list.filter((o): o is SheetShape => !isChart(o) && o.kind === "shape")
const drawing = async (bytes: Uint8Array) =>
  new TextDecoder().decode((await readZip(bytes)).find((e) => e.normalizedName === "xl/drawings/drawing1.xml")!.data)
/** The cNvPr id written for a shape of that name. */
const idOf = (xml: string, name: string) => new RegExp(`<xdr:cNvPr id="(\\d+)" name="${name}"`).exec(xml)?.[1]

describe("connectors glued to shapes in xlsx", () => {
  it("reads Excel's glued ends as Cells' shapes and sites", async () => {
    const { workbook } = await openXlsx(fixture(), "connector.xlsx")
    const [rect, oval, arrow] = shapes(workbook.sheets[0].objects)
    expect(arrow.connector).toBe(true)
    expect(arrow.from).toEqual({ id: rect.id, site: 3 })
    expect(arrow.to).toEqual({ id: oval.id, site: 2 })
  })

  it("writes them naming the shapes by the ids they're written with, changed or not", async () => {
    const { workbook, source } = await openXlsx(fixture(), "connector.xlsx")
    let xml = await drawing(await saveXlsx(workbook, { source, fileName: "c.xlsx" }))
    expect(xml).toContain(`<a:stCxn id="${idOf(xml, "Rectangle 1")}" idx="3"/>`)
    expect(xml).toContain(`<a:endCxn id="${idOf(xml, "Oval 2")}" idx="2"/>`)
    // The oval renamed (written anew, with a new id): the arrow follows.
    const oval = shapes(workbook.sheets[0].objects)[1]
    oval.name = "Target"
    const saved = await saveXlsx(workbook, { source, fileName: "c.xlsx" })
    xml = await drawing(saved)
    expect(xml).toContain(`<a:endCxn id="${idOf(xml, "Target")}" idx="2"/>`)
    const back = shapes((await openXlsx(saved, "c.xlsx")).workbook.sheets[0].objects)
    expect(back[2].to).toEqual({ id: back[1].id, site: 2 })
  })

  it("an end whose shape is gone lets go when saved", async () => {
    const { workbook, source } = await openXlsx(fixture(), "connector.xlsx")
    const sheet = workbook.sheets[0]
    sheet.objects = sheet.objects.filter((o) => o.name !== "Oval 2")
    const xml = await drawing(await saveXlsx(workbook, { source, fileName: "c.xlsx" }))
    expect(xml).toContain("<a:stCxn")
    expect(xml).not.toContain("<a:endCxn")
  })
})
