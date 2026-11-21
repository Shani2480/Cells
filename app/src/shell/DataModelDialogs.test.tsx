import { readFileSync } from "node:fs"
import { act } from "react"
import { createRoot } from "react-dom/client"
import { afterEach, describe, expect, it } from "vitest"
import { openXlsx } from "@cells/io"
import { SheetController } from "@/sheet/controller"
import { DataModelDialog, modelColumns, RelationshipsDialog } from "./DataModelDialogs"

/** Excel's workbook (mock data): Orders and Products in the Data Model, related on ProductID. */
async function modelWorkbook() {
  const bytes = new Uint8Array(readFileSync("packages/io/test/fixtures/excel/data-model.xlsx"))
  return new SheetController((await openXlsx(bytes, "data-model.xlsx")).workbook)
}

let cleanup: (() => void) | null = null
afterEach(() => {
  cleanup?.()
  cleanup = null
  document.body.innerHTML = ""
})

async function show(node: React.ReactNode) {
  const host = document.createElement("div")
  document.body.append(host)
  const root = createRoot(host)
  await act(async () => root.render(node))
  cleanup = () => root.unmount()
}

const button = (t: string) => [...document.body.querySelectorAll("button")].find((b) => b.textContent === t)!

describe("Data Model dialogs", () => {
  it("Manage Relationships lists Excel's relationship, with changes left to Excel", async () => {
    const c = await modelWorkbook()
    let managed = false
    await show(<RelationshipsDialog controller={c} onClose={() => {}} onManage={() => (managed = true)} />)
    const rows = [...document.body.querySelectorAll('button[role="row"]')].map((r) =>
      [...r.querySelectorAll('[role="gridcell"]')].map((x) => x.textContent),
    )
    expect(rows).toEqual([["Active", "Orders (ProductID)", "Products (ProductID)"]])
    expect(button("New...").disabled).toBe(true)
    expect(button("Delete").disabled).toBe(true)
    await act(async () => button("Data Model...").click())
    expect(managed).toBe(true)
  })

  it("Manage Data Model draws the tables and the relationship, and shows a table's data", async () => {
    const c = await modelWorkbook()
    expect(modelColumns(c.wb, c.wb.dataModel!, "Products")).toEqual(["ProductID", "Product", "Price"])
    await show(<DataModelDialog controller={c} onClose={() => {}} />)
    const titles = [...document.body.querySelectorAll(".dm-box-title")].map((t) => t.textContent)
    expect(titles).toEqual(["Orders", "Products"])
    expect(document.body.querySelectorAll(".dm-link path")).toHaveLength(1)
    expect([...document.body.querySelectorAll(".dm-box-col.is-key")].map((t) => t.textContent)).toEqual([
      "ProductID",
      "ProductID",
    ])
    await act(async () => button("Data View").click())
    await act(async () => button("Products").click())
    const cells = [...document.body.querySelectorAll(".dm-grid tbody tr")].map((r) =>
      [...r.querySelectorAll("td")].map((d) => d.textContent),
    )
    expect(cells).toEqual([
      ["10", "Pens", "2.5"],
      ["11", "Ink", "4"],
      ["12", "Paper", "1.25"],
    ])
    expect(document.body.querySelector(".dm-count")?.textContent).toBe("3 rows")
  })

  it("explains a workbook without a model", async () => {
    await show(<DataModelDialog controller={new SheetController()} onClose={() => {}} />)
    expect(document.body.querySelector(".dm-empty")?.textContent).toContain("This workbook has no Data Model")
  })
})
