import { writeFileSync } from "node:fs"
import { describe, expect, it, vi } from "vitest"
import { parseInput, readQueryPlan, selectRange, type QueryPlan } from "@cells/core"
import { openXlsx, saveXlsx } from "@cells/io"
import { nameFromUrl, webKind } from "./get-data"
import { SheetController } from "./controller"
import {
  detectTypes,
  htmlTables,
  jsonTable,
  loadOps,
  parseCsv,
  refreshTableOps,
  runPlan,
  type QueryContext,
} from "./power-query"

const SALES_NORTH = `let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Region", type text}, {"Product", type text}, {"Units", Int64.Type}, {"Price", type number}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type",{"Region", "Product", "Units"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Region] = "North")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Units", Order.Descending}}),
    #"Removed Duplicates" = Table.Distinct(#"Sorted Rows")
in
    #"Removed Duplicates"`

/** A workbook with Excel's example's mock table "Sales" in A1:D6. */
function sales() {
  const c = new SheetController()
  const rows = [
    ["Region", "Product", "Units", "Price"],
    ["North", "Pens", "12", "2.5"],
    ["South", "Ink", "5", "4"],
    ["North", "Paper", "40", "1.25"],
    ["East", "Pens", "7", "2.5"],
    ["North", "Pens", "12", "2.5"],
  ]
  const cells = rows.flatMap((r, row) =>
    r.map((t, col) => {
      const p = parseInput(t)!
      return { row, col, data: { value: p.value } }
    }),
  )
  c.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" })
  c.tables.create({ r0: 0, c0: 0, r1: 5, c1: 3 }, true, "TableStyleMedium2")
  c.wb.transact([{ kind: "tables", sheet: 0, tables: [{ ...c.sheet.tables[0], name: "Sales" }] }], { label: "Name" })
  return c
}

const offline: (c: SheetController) => QueryContext = (c) => ({
  wb: c.wb,
  fetch: async () => {
    throw new Error("offline")
  },
})

describe("running queries", () => {
  it("gives what Excel's refresh gave for its own query", async () => {
    const c = sales()
    const r = await runPlan(readQueryPlan(SALES_NORTH)!, offline(c))
    expect(r.columns).toEqual(["Region", "Product", "Units"])
    expect(r.rows).toEqual([
      ["North", "Paper", 40],
      ["North", "Pens", 12],
    ])
  })

  it("loads into a table on a sheet named after the query, and refresh follows the source", async () => {
    const c = sales()
    const plan = readQueryPlan(SALES_NORTH)!
    const query = { name: "Sales North", description: "", formula: SALES_NORTH, table: null }
    const { ops, sheet } = loadOps(c.wb, query, await runPlan(plan, offline(c)), 0)
    c.wb.transact(ops, { label: "Load" })
    expect(c.wb.sheets[sheet].name).toBe("Sales North")
    const t = c.wb.sheets[sheet].tables[0]
    expect(t).toMatchObject({ name: "Sales_North", query: "Sales North", range: { r0: 0, c0: 0, r1: 2, c1: 2 } })
    expect(c.wb.queries[0].table).toBe("Sales_North")
    // A third north row in the source: refresh grows the table.
    c.wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 2, col: 0, data: { value: "North" } }] }], {
      label: "Edit",
    })
    const again = await runPlan(plan, offline(c))
    c.wb.transact(refreshTableOps(c.wb, c.wb.queries[0], again), { label: "Refresh" })
    expect(c.wb.sheets[sheet].tables[0].range.r1).toBe(3)
    expect(c.wb.sheets[sheet].cells.value(3, 1)).toBe("Ink")
    // Fewer rows: the table shrinks and the old row is cleared.
    c.wb.transact(
      [
        {
          kind: "setCells",
          sheet: 0,
          cells: [
            { row: 3, col: 0, data: { value: "West" } },
            { row: 2, col: 0, data: { value: "South" } },
          ],
        },
      ],
      { label: "Edit" },
    )
    c.wb.transact(refreshTableOps(c.wb, c.wb.queries[0], await runPlan(plan, offline(c))), { label: "Refresh" })
    expect(c.wb.sheets[sheet].tables[0].range.r1).toBe(1)
    expect(c.wb.sheets[sheet].cells.value(2, 0)).toBeUndefined()
  })

  it("reads CSV, a web page's tables and JSON from the web", async () => {
    const files: Record<string, string> = {
      "https://example.com/prices.csv": 'Item,Price,Note\r\nPens,2.5,"blue, black"\r\nInk,4,\r\n',
      "https://example.com/page":
        "<table><tr><th>City</th><th>Rate</th></tr><tr><td>Lisbon</td><td>1.5</td></tr></table><table><tr><td>x</td></tr></table>",
      "https://example.com/a.json": '{"items":[{"id":1,"name":"Pens"},{"id":2,"name":"Ink","tag":"new"}]}',
    }
    const ctx: QueryContext = {
      wb: new SheetController().wb,
      fetch: async (url) => ({ bytes: new TextEncoder().encode(files[url]), type: "" }),
    }
    const csv: QueryPlan = {
      source: { kind: "csv", url: "https://example.com/prices.csv", delimiter: ",", encoding: 65001 },
      steps: [{ name: "Promoted Headers", step: { op: "promote" } }],
    }
    const r = await runPlan(csv, ctx)
    expect(r.columns).toEqual(["Item", "Price", "Note"])
    expect(r.rows).toEqual([
      ["Pens", "2.5", "blue, black"],
      ["Ink", "4", null],
    ])
    expect(detectTypes(r)).toEqual([
      ["Item", "text"],
      ["Price", "number"],
      ["Note", "text"],
    ])
    const typed = await runPlan(
      { ...csv, steps: [...csv.steps, { name: "Changed Type", step: { op: "types", columns: detectTypes(r) } }] },
      ctx,
    )
    expect(typed.rows[0][1]).toBe(2.5)
    const page = await runPlan(
      {
        source: { kind: "html", url: "https://example.com/page" },
        steps: [{ name: "Data0", step: { op: "table", index: 0 } }],
      },
      ctx,
    )
    expect(page).toMatchObject({ columns: ["City", "Rate"], rows: [["Lisbon", "1.5"]] })
    const json = await runPlan({ source: { kind: "json", url: "https://example.com/a.json" }, steps: [] }, ctx)
    expect(json.columns).toEqual(["id", "name", "tag"])
    expect(json.rows[0]).toEqual([1, "Pens", null])
  })

  it("parses CSV quotes and line ends, and says what's wrong", async () => {
    expect(parseCsv('a;"b ""c"""\nd;e', ";")).toEqual([
      ["a", 'b "c"'],
      ["d", "e"],
    ])
    expect(htmlTables("<p>none</p>")).toEqual([])
    expect(() => jsonTable("[1,2]")).toThrow(/list of records/)
    const c = sales()
    const bad: QueryPlan = {
      source: { kind: "table", name: "Sales" },
      steps: [{ name: "Filtered Rows", step: { op: "filter", column: "Colour", test: "=", value: "x" } }],
    }
    await expect(runPlan(bad, offline(c))).rejects.toThrow("The column 'Colour' of the table wasn't found.")
  })
})

describe("Get Data", () => {
  it("From Table/Range, a filter, then Close & Load; undo takes it all back", async () => {
    const c = sales()
    c.setSelection(selectRange({ r0: 1, c0: 1, r1: 1, c1: 1 }))
    expect(c.getData.fromTable()).toBe(true)
    await vi.waitFor(() => expect(c.getData.editor).not.toBeNull())
    const request = c.getData.editor!
    expect(request.query.name).toBe("Sales")
    expect(request.plan.steps.map((s) => s.name)).toEqual(["Changed Type"])
    request.plan.steps.push({ name: "Filtered Rows", step: { op: "filter", column: "Units", test: ">", value: 10 } })
    await c.getData.load(request, true)
    expect(c.wb.sheets[1].name).toBe("Sales")
    expect(c.sheetIndex).toBe(1)
    const t = c.sheet.tables[0]
    expect(t).toMatchObject({ name: "Sales_2", query: "Sales", range: { r1: 3, c1: 3 } })
    expect(c.wb.queries[0]).toMatchObject({ name: "Sales", table: "Sales_2" })
    expect(c.wb.queries[0].formula).toMatch(/Table.SelectRows\(#"Changed Type", each \(?\[Units\] > 10/)
    // Edit, then load into the same table again.
    expect(c.getData.editQuery("Sales")).toBe(true)
    const again = c.getData.editor!
    again.plan.steps.push({ name: "Kept First Rows", step: { op: "first", count: 1 } })
    await c.getData.load(again, true)
    expect(c.wb.sheets.length).toBe(2)
    expect(c.sheet.tables[0].range.r1).toBe(1)
    c.undo()
    c.undo()
    expect(c.wb.queries).toEqual([])
    expect(c.wb.sheets.length).toBe(1)
  })

  it("keeps a connection-only query, and Delete leaves the table as plain data", async () => {
    const c = sales()
    c.getData.openTable("Sales")
    await vi.waitFor(() => expect(c.getData.editor).not.toBeNull())
    await c.getData.load(c.getData.editor!, false)
    expect(c.wb.queries).toMatchObject([{ name: "Sales", table: null }])
    expect(c.wb.sheets.length).toBe(1)
    await c.getData.loadConnection("Sales")
    expect(c.wb.queries[0].table).toBe("Sales_2")
    c.getData.remove("Sales")
    expect(c.wb.queries).toEqual([])
    expect(c.wb.sheets[1].tables[0].query).toBeUndefined()
  })

  it("tells CSV, JSON and pages apart, and names queries after the address", () => {
    const b = (s: string) => new TextEncoder().encode(s)
    expect(webKind("https://example.com/x", "application/json; charset=utf-8", b(""))).toBe("json")
    expect(webKind("https://example.com/data.csv?x=1", "application/octet-stream", b("a,b"))).toBe("csv")
    expect(webKind("https://example.com/", "", b("  <!doctype html>"))).toBe("html")
    expect(nameFromUrl("https://example.com/files/Monthly%20Sales.csv")).toBe("Monthly Sales")
    expect(nameFromUrl("https://example.com/")).toBe("example.com")
  })

  it("saves a query Cells made into a file that reads back the same", async () => {
    const c = sales()
    c.getData.openTable("Sales")
    await vi.waitFor(() => expect(c.getData.editor).not.toBeNull())
    const request = c.getData.editor!
    request.plan.steps.push({ name: "Sorted Rows", step: { op: "sort", keys: [{ column: "Units", desc: true }] } })
    await c.getData.load(
      { ...request, query: { ...request.query, name: "Top Sales", description: "Biggest first" } },
      true,
    )
    const saved = await saveXlsx(c.wb)
    if (process.env.PQ_OUT) writeFileSync(process.env.PQ_OUT, saved)
    const { workbook } = await openXlsx(saved, "x.xlsx")
    expect(workbook.queries).toEqual([{ ...c.wb.queries[0] }])
    expect(workbook.sheets[1].tables[0]).toMatchObject({ name: "Top_Sales", query: "Top Sales" })
  })
})
