import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { FindReplaceDialog, GoToSpecialDialog } from "./FindDialogs"

function sheet() {
  const c = new SheetController()
  c.sheet.cells.set(0, 0, { value: "Apple" })
  c.sheet.cells.set(3, 2, { value: "apple pie" })
  c.sheet.cells.set(5, 1, { value: 7 })
  return c
}

describe("Find and Replace dialog", () => {
  it("Find Next, Options and Find All", () => {
    const c = sheet()
    render(<FindReplaceDialog controller={c} tab="find" onClose={() => {}} />)
    expect(screen.getByRole("dialog")).not.toHaveAttribute("aria-modal")
    expect(screen.getByRole("button", { name: "Find Next" })).toBeDisabled()
    fireEvent.change(screen.getByRole("textbox", { name: "Find what" }), { target: { value: "apple" } })
    fireEvent.click(screen.getByRole("button", { name: "Find Next" }))
    expect(c.selection.active).toEqual({ row: 3, col: 2 })
    fireEvent.click(screen.getByRole("button", { name: /Options/ }))
    fireEvent.click(screen.getByRole("checkbox", { name: /Match case/ }))
    fireEvent.click(screen.getByRole("button", { name: "Find All" }))
    expect(screen.getByText("1 cell found")).toBeInTheDocument()
    expect(screen.getByRole("cell", { name: "$C$4" })).toBeInTheDocument()
  })

  it("the Replace tab replaces everything", () => {
    const c = sheet()
    c.shell.alert = vi.fn()
    render(<FindReplaceDialog controller={c} tab="find" onClose={() => {}} />)
    fireEvent.click(screen.getByRole("tab", { name: "Replace" }))
    fireEvent.change(screen.getByRole("textbox", { name: "Find what" }), { target: { value: "apple" } })
    fireEvent.change(screen.getByRole("textbox", { name: "Replace with" }), { target: { value: "Pear" } })
    fireEvent.click(screen.getByRole("button", { name: "Replace All" }))
    expect(c.sheet.cells.get(0, 0)?.value).toBe("Pear")
    expect(c.sheet.cells.get(3, 2)?.value).toBe("Pear pie")
    expect(c.shell.alert).toHaveBeenCalledWith("All done. We made 2 replacements.")
  })
})

describe("Go To Special dialog", () => {
  it("Constants without text selects the numbers", () => {
    const c = sheet()
    const onClose = vi.fn()
    render(<GoToSpecialDialog controller={c} onClose={onClose} />)
    fireEvent.click(screen.getByRole("checkbox", { name: "Text" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.selection.ranges).toEqual([{ r0: 5, c0: 1, r1: 5, c1: 1 }])
  })
})
