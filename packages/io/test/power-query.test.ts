import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { queryFormula, readQueryPlan, type Workbook } from "@cells/core"
import { openXlsx, saveXlsx } from "../src"
import { readZip } from "@cells/ooxml"

/**
 * Excel's file of Power Query queries (mock data): "Sales North" reads the table Sales, keeps three columns, filters
 * to the north, sorts and removes duplicates, and loads into Table_ExternalData_1 on Sheet2; "Prices" reads a CSV
 * from the web and is a connection only.
 */
const bytes = new Uint8Array(readFileSync(new URL("./fixtures/excel/power-query.xlsx", import.meta.url)))

const text = (files: { name: string; data: Uint8Array }[], name: string) =>
  new TextDecoder().decode(files.find((f) => f.name === name)?.data ?? new Uint8Array())

async function reopen(wb: Workbook, source?: unknown) {
  const saved = await saveXlsx(wb, source ? { source: source as never } : {})
  return { saved, again: await openXlsx(saved, "again.xlsx"), files: await readZip(saved) }
}

describe("Power Query in xlsx", () => {
  it("reads Excel's queries, their descriptions and the table each loads", async () => {
    const { workbook: wb } = await openXlsx(bytes, "power-query.xlsx")
    expect(wb.queries.map((q) => [q.name, q.description, q.table])).toEqual([
      ["Sales North", "Units of the north", "Table_ExternalData_1"],
      ["Prices", "", null],
    ])
    const plan = readQueryPlan(wb.queries[0].formula)!
    expect(plan.source).toEqual({ kind: "table", name: "Sales" })
    expect(plan.steps.map((s) => s.step.op)).toEqual(["types", "select", "filter", "sort", "distinct"])
    expect(queryFormula(plan)).toBe(wb.queries[0].formula)
    const loaded = wb.sheets[1].tables[0]
    expect(loaded.query).toBe("Sales North")
    // The hidden ExternalData name is the file's business, not a name people see.
    expect(wb.names.all().some((n) => /^ExternalData/.test(n.name))).toBe(false)
  })

  it("saves untouched queries as Excel wrote them, with the query table part kept", async () => {
    const r = await openXlsx(bytes, "power-query.xlsx")
    const { files, again } = await reopen(r.workbook, r.source)
    const item = files.find(
      (f) => /^customXml\/item\d+\.xml$/.test(f.name) && /DataMashup/.test(new TextDecoder("utf-16le").decode(f.data)),
    )
    const original = (await readZip(bytes)).find((f) => f.name === "customXml/item1.xml")!
    expect(item?.data).toEqual(original.data)
    const queryTable = files.find((f) => /^xl\/queryTables\/queryTable\d+\.xml$/.test(f.name))
    expect(queryTable).toBeTruthy()
    const table = files.find((f) => /tables\/table\d+\.xml$/.test(f.name) && /queryTable/.test(text(files, f.name)))
    expect(text(files, table!.name)).toMatch(/tableType="queryTable"/)
    expect(text(files, "xl/connections.xml")).toMatch(
      /Microsoft\.Mashup\.OleDb\.1.*Location=&quot;?Sales North|Location=Sales North/,
    )
    expect(again.workbook.queries.map((q) => q.table)).toEqual(["Table_ExternalData_1", null])
    expect(text(files, "xl/workbook.xml")).toMatch(
      /name="ExternalData_1"[^>]*hidden="1"|hidden="1"[^>]*name="ExternalData_1"/,
    )
  })

  it("writes a changed query anew, and a workbook's first query from nothing", async () => {
    const r = await openXlsx(bytes, "power-query.xlsx")
    const q = r.workbook.queries[0]
    const plan = readQueryPlan(q.formula)!
    plan.steps = plan.steps.map((s) => (s.step.op === "filter" ? { ...s, step: { ...s.step, value: "South" } } : s))
    r.workbook.transact(
      [
        {
          kind: "queries",
          queries: [{ ...q, formula: queryFormula(plan), description: "Units of the south" }, r.workbook.queries[1]],
        },
      ],
      { label: "Edit" },
    )
    const { again } = await reopen(r.workbook, r.source)
    expect(again.workbook.queries[0].description).toBe("Units of the south")
    expect(again.workbook.queries[0].formula).toContain('[Region] = "South"')
    expect(again.workbook.queries[1].formula).toContain("Csv.Document")

    // Deleting every query takes the DataMashup item with it.
    r.workbook.transact([{ kind: "queries", queries: [] }], { label: "Delete" })
    const gone = await reopen(r.workbook, r.source)
    expect(gone.files.some((f) => /DataMashup/.test(new TextDecoder("utf-16le").decode(f.data)))).toBe(false)
    expect(gone.again.workbook.queries).toEqual([])
  })
})
