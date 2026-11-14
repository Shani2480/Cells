import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { InsertFunctionDialog, searchFunctions } from "./InsertFunctionDialog"

describe("Insert Function", () => {
  it("lists a category's functions with the picked one's signature, and OK inserts it", () => {
    const onInsert = vi.fn()
    render(<InsertFunctionDialog onInsert={onInsert} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText(/category/), { target: { value: "Logical" } })
    const options = screen.getAllByRole("option").map((o) => o.textContent)
    expect(options).toContain("IFERROR")
    expect(options).not.toContain("SUM")
    fireEvent.click(screen.getByRole("option", { name: "IFERROR" }))
    expect(screen.getByText("IFERROR(value,value_if_error)")).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onInsert).toHaveBeenCalledWith("IFERROR")
  })

  it("searches names and descriptions", () => {
    expect(searchFunctions("average")).toContain("AVERAGEIF")
    expect(searchFunctions("internal rate")).toEqual(expect.arrayContaining(["IRR", "XIRR"]))
    expect(searchFunctions("   ")).toEqual([])
  })
})
