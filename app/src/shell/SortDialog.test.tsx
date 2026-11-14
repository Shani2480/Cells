import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { SortDialog } from "./SortDialog"

function table() {
  const c = new SheetController()
  const rows: (string | number)[][] = [
    ["Name", "Team", "Month"],
    ["Ann", "Blue", "Mar"],
    ["Bob", "Red", "Jan"],
    ["Cy", "Blue", "Feb"],
  ]
  rows.forEach((row, r) => row.forEach((value, col) => c.sheet.cells.set(r, col, { value })))
  const bold = c.wb.styles.intern({ bold: true })
  for (let col = 0; col < 3; col++) c.sheet.cells.set(0, col, { value: rows[0][col], style: bold })
  return c
}

describe("Sort dialog", () => {
  it("names columns by their headers and sorts by several levels", () => {
    const c = table()
    const onClose = vi.fn()
    render(<SortDialog controller={c} range={{ r0: 0, c0: 0, r1: 3, c1: 2 }} onClose={onClose} />)
    expect(screen.getByRole("checkbox", { name: "My data has headers" })).toBeChecked()
    const first = screen.getByRole("combobox", { name: "Sort by column" })
    expect([...(first as HTMLSelectElement).options].map((o) => o.text)).toEqual(["Name", "Team", "Month"])
    fireEvent.change(first, { target: { value: "1" } })
    fireEvent.click(screen.getByRole("button", { name: /Add Level/ }))
    const then = screen.getByRole("combobox", { name: "Then by column" })
    fireEvent.change(then, { target: { value: "0" } })
    fireEvent.change(screen.getAllByRole("combobox", { name: "Order" })[1], { target: { value: "desc" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect([0, 1, 2, 3].map((r) => c.sheet.cells.get(r, 0)?.value)).toEqual(["Name", "Cy", "Ann", "Bob"])
  })

  it("a custom list orders months as in the calendar; levels move and delete", () => {
    const c = table()
    render(<SortDialog controller={c} range={{ r0: 0, c0: 0, r1: 3, c1: 2 }} onClose={() => {}} />)
    fireEvent.change(screen.getByRole("combobox", { name: "Sort by column" }), { target: { value: "2" } })
    const order = screen.getByRole("combobox", { name: "Order" }) as HTMLSelectElement
    const months = [...order.options].find((o) => o.text.startsWith("Jan, Feb"))!
    fireEvent.change(order, { target: { value: months.value } })
    fireEvent.click(screen.getByRole("button", { name: /Add Level/ }))
    expect(screen.getAllByRole("row")).toHaveLength(3)
    fireEvent.click(screen.getByRole("button", { name: /Delete Level/ }))
    expect(screen.getAllByRole("row")).toHaveLength(2)
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect([1, 2, 3].map((r) => c.sheet.cells.get(r, 2)?.value)).toEqual(["Jan", "Feb", "Mar"])
  })

  it("Options: Enter in the options dialog doesn't also sort", () => {
    const c = table()
    const onClose = vi.fn()
    render(<SortDialog controller={c} range={{ r0: 0, c0: 0, r1: 3, c1: 2 }} onClose={onClose} />)
    fireEvent.click(screen.getByRole("button", { name: "Options..." }))
    fireEvent.click(screen.getByRole("checkbox", { name: "Case sensitive" }))
    fireEvent.keyDown(screen.getByRole("checkbox", { name: "Case sensitive" }), { key: "Enter" })
    expect(screen.queryByText("Sort Options")).toBeNull()
    expect(onClose).not.toHaveBeenCalled()
  })
})
