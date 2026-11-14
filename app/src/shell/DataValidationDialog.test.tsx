import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { DEFAULT_VALIDATION } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { DataValidationDialog, ValidationAlertDialog } from "./DataValidationDialog"

const A1_A5 = [{ r0: 0, c0: 0, r1: 4, c1: 0 }]

describe("Data Validation dialog", () => {
  it("sets a whole-number rule with an input message and a warning", () => {
    const c = new SheetController()
    const onClose = vi.fn()
    render(<DataValidationDialog controller={c} init={{ rule: DEFAULT_VALIDATION, ranges: A1_A5 }} onClose={onClose} />)
    fireEvent.change(screen.getByLabelText("Allow:"), { target: { value: "whole" } })
    fireEvent.change(screen.getByLabelText("Minimum:"), { target: { value: "1" } })
    fireEvent.change(screen.getByLabelText("Maximum:"), { target: { value: "10" } })
    fireEvent.click(screen.getByRole("tab", { name: "Input Message" }))
    fireEvent.change(screen.getByLabelText("Title:"), { target: { value: "Qty" } })
    fireEvent.change(screen.getByLabelText("Input message:"), { target: { value: "1 to 10" } })
    fireEvent.click(screen.getByRole("tab", { name: "Error Alert" }))
    fireEvent.change(screen.getByLabelText("Style:"), { target: { value: "warning" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.validation.at(4, 0)?.rule).toMatchObject({
      type: "whole",
      formula1: "1",
      formula2: "10",
      promptTitle: "Qty",
      prompt: "1 to 10",
      errorStyle: "warning",
    })
  })

  it("refuses a Maximum below the Minimum, and a list without a source", () => {
    const c = new SheetController()
    c.shell.alert = vi.fn()
    const onClose = vi.fn()
    render(<DataValidationDialog controller={c} init={{ rule: DEFAULT_VALIDATION, ranges: A1_A5 }} onClose={onClose} />)
    fireEvent.change(screen.getByLabelText("Allow:"), { target: { value: "decimal" } })
    fireEvent.change(screen.getByLabelText("Minimum:"), { target: { value: "5" } })
    fireEvent.change(screen.getByLabelText("Maximum:"), { target: { value: "2" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.shell.alert).toHaveBeenLastCalledWith("The Maximum must be greater than or equal to the Minimum.")
    fireEvent.change(screen.getByLabelText("Allow:"), { target: { value: "list" } })
    expect(screen.getByRole("checkbox", { name: "In-cell dropdown" })).toBeChecked()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.shell.alert).toHaveBeenLastCalledWith(
      "The list source must be a delimited list, or a reference to single row or column.",
    )
    expect(onClose).not.toHaveBeenCalled()
  })

  it("opens on a rule's values; Clear All and OK remove it", () => {
    const c = new SheetController()
    const rule = { ...DEFAULT_VALIDATION, type: "list" as const, formula1: '"Yes,No"' }
    c.validation.apply(A1_A5, rule)
    render(<DataValidationDialog controller={c} init={{ rule, ranges: A1_A5 }} onClose={() => {}} />)
    expect(screen.getByLabelText("Source:")).toHaveValue("Yes,No")
    expect(screen.getByRole("checkbox", { name: /same settings/ })).toBeEnabled()
    fireEvent.click(screen.getByRole("button", { name: "Clear All" }))
    expect(screen.getByLabelText("Allow:")).toHaveValue("any")
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.validations).toEqual([])
  })
})

describe("error alert", () => {
  it("Warning asks Continue? with Yes, No and Cancel", () => {
    const onAnswer = vi.fn()
    render(
      <ValidationAlertDialog alert={{ style: "warning", title: "Check", message: "Too big" }} onAnswer={onAnswer} />,
    )
    expect(screen.getByRole("alert")).toHaveTextContent("Too bigContinue?")
    fireEvent.click(screen.getByRole("button", { name: "No" }))
    expect(onAnswer).toHaveBeenCalledWith("no")
  })
})
