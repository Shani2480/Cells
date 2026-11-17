import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import type { FileCommands } from "@/file/commands"
import { SheetController } from "@/sheet/controller"
import { QueriesPane } from "./QueriesPane"
import { QueryPropertiesDialog } from "./QueryPropertiesDialog"

const OPTIONS = {
  kind: "delimited" as const,
  delimiters: [","],
  qualifier: '"',
  consecutive: false,
  breaks: [],
  startRow: 1,
  columns: [],
  decimal: ".",
  thousands: ",",
}

function setup() {
  const c = new SheetController()
  c.transact(
    [
      {
        kind: "queryTables",
        sheet: 0,
        tables: [
          {
            name: "sales",
            connection: "sales",
            description: "",
            range: { r0: 0, c0: 0, r1: 4, c1: 2 },
            source: "/home/alex/sales.csv",
            options: OPTIONS,
            refreshOnLoad: false,
          },
        ],
      },
    ],
    { label: "Link", before: c.selection },
  )
  c.data.toggleQueries(true)
  return c
}

describe("Queries & Connections pane", () => {
  it("lists each link with its rows, file and range; Refresh, Properties and Delete act on it", () => {
    const c = setup()
    const refreshQuery = vi.fn(async () => true)
    const refreshAll = vi.fn(async () => 1)
    const files = { refreshQuery, refreshAll } as unknown as FileCommands
    const onProperties = vi.fn()
    render(<QueriesPane controller={c} files={files} onProperties={onProperties} onDone={() => {}} />)
    expect(screen.getByText("5 rows loaded · sales.csv")).toBeTruthy()
    fireEvent.click(screen.getByRole("button", { name: "sales, Sheet1!A1:C5" }))
    expect(c.selection.ranges[0]).toEqual({ r0: 0, c0: 0, r1: 4, c1: 2 })
    fireEvent.click(screen.getByRole("button", { name: "Refresh sales" }))
    expect(refreshQuery).toHaveBeenCalledWith(0, "sales")
    fireEvent.click(screen.getByRole("button", { name: "Properties of sales" }))
    expect(onProperties).toHaveBeenCalled()
    expect(c.data.queryTarget).toEqual({ sheet: 0, name: "sales" })
    fireEvent.click(screen.getByRole("button", { name: "Delete sales" }))
    expect(c.sheet.queryTables).toEqual([])
    expect(screen.getByText(/This workbook has no connections/)).toBeTruthy()
  })

  it("Connection Properties renames, describes and sets refresh on open; bad names are refused", () => {
    const c = setup()
    c.data.queryTarget = { sheet: 0, name: "sales" }
    const onClose = vi.fn()
    render(<QueryPropertiesDialog controller={c} onClose={onClose} />)
    fireEvent.change(screen.getByLabelText("Data range name:"), { target: { value: "A1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(screen.getByRole("alert").textContent).toMatch(/not valid/)
    expect(onClose).not.toHaveBeenCalled()
    fireEvent.change(screen.getByLabelText("Data range name:"), { target: { value: "Sales_2026" } })
    fireEvent.change(screen.getByLabelText("Connection name:"), { target: { value: "Monthly sales" } })
    fireEvent.change(screen.getByLabelText("Description:"), { target: { value: "From the shop" } })
    fireEvent.click(screen.getByLabelText("Refresh data when opening the file"))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.sheet.queryTables[0]).toMatchObject({
      name: "Sales_2026",
      connection: "Monthly sales",
      description: "From the shop",
      refreshOnLoad: true,
    })
    c.undo()
    expect(c.sheet.queryTables[0].name).toBe("sales")
  })
})
