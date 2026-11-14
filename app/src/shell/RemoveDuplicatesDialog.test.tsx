import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { RemoveDuplicatesDialog } from "./RemoveDuplicatesDialog"

function sheet() {
  const c = new SheetController()
  const rows = [
    ["Name", "Age"],
    ["Ann", 30],
    ["Bob", 40],
    ["Ann", 30],
    ["Ann", 40],
  ]
  rows.forEach((r, i) => r.forEach((v, j) => c.sheet.cells.set(i, j, { value: v, input: String(v) })))
  c.shell.alert = vi.fn()
  return c
}
const range = { r0: 0, c0: 0, r1: 4, c1: 1 }

describe("Remove Duplicates dialog", () => {
  it("lists the headers and removes rows that repeat in every column", () => {
    const c = sheet()
    const onClose = vi.fn()
    render(<RemoveDuplicatesDialog controller={c} range={range} onClose={onClose} />)
    expect(screen.getByRole("checkbox", { name: /My data has headers/ })).toBeChecked()
    expect(screen.getByRole("checkbox", { name: "Name" })).toBeChecked()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.sheet.cells.get(3, 1)?.value).toBe(40)
    expect(c.sheet.cells.get(4, 0)).toBeUndefined()
    expect(c.shell.alert).toHaveBeenCalledWith("1 duplicate value found and removed; 3 unique values remain.")
  })

  it("only the ticked columns count, and without headers they are named by letter", () => {
    const c = sheet()
    render(<RemoveDuplicatesDialog controller={c} range={range} onClose={() => {}} />)
    fireEvent.click(screen.getByRole("checkbox", { name: /My data has headers/ }))
    expect(screen.getByRole("checkbox", { name: "Column A" })).toBeInTheDocument()
    fireEvent.click(screen.getByRole("checkbox", { name: /My data has headers/ }))
    fireEvent.click(screen.getByRole("button", { name: "Unselect All" }))
    fireEvent.click(screen.getByRole("checkbox", { name: "Name" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect([1, 2, 3, 4].map((r) => c.sheet.cells.get(r, 0)?.value)).toEqual(["Ann", "Bob", undefined, undefined])
  })
})
