import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

/**
 * Excel's file with a Data Model (mock data): tables Orders and Products added to the model through worksheet
 * connections, related on ProductID. Excel keeps the model in xl/model/item.data and trusts only that part.
 */
const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/data-model.xlsx", import.meta.url)))

describe("Data Model in xlsx", () => {
  it("keeps Excel's model part, its relationship and its connections well-formed", async () => {
    const r = await openXlsx(bytes, "data-model.xlsx")
    const saved = await saveXlsx(r.workbook, { source: r.source })
    const files = await readZip(saved)
    const text = (name: string) => new TextDecoder().decode(files.find((f) => f.name === name)!.data)
    const original = (await readZip(bytes)).find((f) => f.name === "xl/model/item.data")!
    expect(files.find((f) => f.name === "xl/model/item.data")?.data).toEqual(original.data)
    expect(text("xl/_rels/workbook.xml.rels")).toMatch(/relationships\/powerPivotData" Target="model\/item.data"/)
    expect(text("xl/workbook.xml")).toContain(
      '<x15:modelRelationship fromTable="Orders" fromColumn="ProductID" toTable="Products" toColumn="ProductID"/>',
    )
    // Excel's connections carry xr16:uid; the root must declare it, or Excel can't open the file at all.
    const connections = text("xl/connections.xml")
    expect(connections).toContain("xr16:uid=")
    expect(connections).toMatch(
      /<connections [^>]*xmlns:xr16="http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2017\/revision16"/,
    )
    expect(connections).toMatch(/mc:Ignorable="[^"]*xr16/)
  })
})

describe("Data Model, read", () => {
  it("lists the model's tables with the workbook tables they load, and the relationships", async () => {
    const { workbook } = await openXlsx(bytes, "data-model.xlsx")
    expect(workbook.dataModel).toEqual({
      tables: [
        { name: "Orders", connection: "WorksheetConnection_Orders", source: "Orders" },
        { name: "Products", connection: "WorksheetConnection_Products", source: "Products" },
      ],
      relationships: [{ fromTable: "Orders", fromColumn: "ProductID", toTable: "Products", toColumn: "ProductID" }],
    })
  })

  it("has no model in a workbook without one", async () => {
    const pq = new Uint8Array(readFileSync(new URL("./fixtures/excel/power-query.xlsx", import.meta.url)))
    expect((await openXlsx(pq, "pq.xlsx")).workbook.dataModel).toBeNull()
  })
})
