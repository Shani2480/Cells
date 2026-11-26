import { describe, expect, it } from "vitest"
import { queryFormula, queryName, readQueryPlan, stepName, Workbook, type QueryPlan } from "../src"

// As Excel 365's Query Editor wrote them (made in the Office VM, with mock data).
const SALES_NORTH = `let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Region", type text}, {"Product", type text}, {"Units", Int64.Type}, {"Price", type number}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type",{"Region", "Product", "Units"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Region] = "North")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Units", Order.Descending}}),
    #"Removed Duplicates" = Table.Distinct(#"Sorted Rows")
in
    #"Removed Duplicates"`
const PRICES = `let
    Source = Csv.Document(Web.Contents("https://example.com/prices.csv"),[Delimiter=",", Columns=3, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    #"Promoted Headers"`

describe("Power Query formulas", () => {
  it("reads the Query Editor's steps", () => {
    const plan = readQueryPlan(SALES_NORTH)!
    expect(plan.source).toEqual({ kind: "table", name: "Sales" })
    expect(plan.steps.map((s) => s.step)).toEqual([
      {
        op: "types",
        columns: [
          ["Region", "text"],
          ["Product", "text"],
          ["Units", "whole"],
          ["Price", "number"],
        ],
      },
      { op: "select", columns: ["Region", "Product", "Units"] },
      { op: "filter", column: "Region", test: "=", value: "North" },
      { op: "sort", keys: [{ column: "Units", desc: true }] },
      { op: "distinct" },
    ])
    expect(readQueryPlan(PRICES)).toEqual({
      source: { kind: "csv", url: "https://example.com/prices.csv", delimiter: ",", encoding: 65001, columns: 3 },
      steps: [{ name: "Promoted Headers", step: { op: "promote" } }],
    })
  })

  it("writes M as the editor does, so Excel's own formula comes back word for word", () => {
    expect(queryFormula(readQueryPlan(SALES_NORTH)!)).toBe(SALES_NORTH.replace(/\n/g, "\r\n"))
    expect(queryFormula(readQueryPlan(PRICES)!)).toBe(PRICES.replace(/\n/g, "\r\n"))
  })

  it("reads and writes web pages, JSON, text tests and kept-first rows", () => {
    const plan: QueryPlan = {
      source: { kind: "html", url: "https://example.com/rates" },
      steps: [
        { name: "Data0", step: { op: "table", index: 0 } },
        { name: "Filtered Rows", step: { op: "filter", column: "Name", test: "contains", value: 'a "b"' } },
        { name: "Filtered Rows1", step: { op: "filter", column: "Rate", test: ">=", value: 1.5 } },
        { name: "Kept First Rows", step: { op: "first", count: 10 } },
        { name: "Removed Columns", step: { op: "remove", columns: ["Notes"] } },
      ],
    }
    expect(readQueryPlan(queryFormula(plan))).toEqual(plan)
    const json: QueryPlan = { source: { kind: "json", url: "https://example.com/a.json" }, steps: [] }
    expect(readQueryPlan(queryFormula(json))).toEqual(json)
  })

  it("reads and writes a PDF's query as Excel's Navigator writes it", () => {
    const excel = `let
    Source = Pdf.Tables(File.Contents("C:\\Users\\alex\\Documents\\statement.pdf"), [Implementation="1.3"]),
    Table001 = Source{[Id="Table001"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Table001, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Amount", type number}})
in
    #"Changed Type"`
    const plan = readQueryPlan(excel)
    expect(plan).toEqual({
      source: { kind: "pdf", path: "C:\\Users\\alex\\Documents\\statement.pdf" },
      steps: [
        { name: "Table001", step: { op: "item", id: "Table001" } },
        { name: "Promoted Headers", step: { op: "promote" } },
        {
          name: "Changed Type",
          step: {
            op: "types",
            columns: [
              ["Date", "date"],
              ["Amount", "number"],
            ],
          },
        },
      ],
    })
    expect(queryFormula(plan as QueryPlan)).toBe(excel.replace(/\n/g, "\r\n"))
    expect(stepName({ op: "item", id: "Page002" }, [])).toBe("Page002")
    // Excel's own options (which change the tables it finds) are left to Excel.
    expect(
      readQueryPlan('let Source = Pdf.Tables(File.Contents("/home/alex/a.pdf"), [EnforceBorderLines=true]) in Source'),
    ).toBeNull()
    expect(readQueryPlan('let Source = Pdf.Tables(File.Contents("/home/alex/a.pdf")) in Source')?.source).toEqual({
      kind: "pdf",
      path: "/home/alex/a.pdf",
    })
  })

  it("leaves M it doesn't run alone", () => {
    expect(readQueryPlan('let Source = Sql.Database("srv", "db") in Source')).toBeNull()
    expect(
      readQueryPlan('let Source = Excel.CurrentWorkbook(){[Name="A"]}[Content], X = Table.Group(Source, {}, {}) in X'),
    ).toBeNull()
    expect(readQueryPlan("not M at all")).toBeNull()
  })

  it("names steps and queries as Excel does", () => {
    expect(stepName({ op: "filter", column: "a", test: "=", value: 1 }, ["Source", "Filtered Rows"])).toBe(
      "Filtered Rows1",
    )
    expect(stepName({ op: "table", index: 2 }, [])).toBe("Data2")
    expect(queryName("Sales", ["sales"])).toBe("Sales (2)")
  })

  it("changes the queries as one step to undo", () => {
    const wb = new Workbook(["Sheet1"])
    wb.transact([{ kind: "queries", queries: [{ name: "Q", description: "", formula: PRICES, table: null }] }], {
      label: "New Query",
    })
    expect(wb.queries.map((q) => q.name)).toEqual(["Q"])
    wb.undo()
    expect(wb.queries).toEqual([])
  })
})
