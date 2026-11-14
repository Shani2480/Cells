import { fireEvent, render, screen } from "@testing-library/react"
import type { TextImport } from "@cells/io"
import { describe, expect, it, vi } from "vitest"
import { TextImportWizard } from "./TextImportWizard"

const bytes = (text: string) => new TextEncoder().encode(text)

function open(text: string, fileName = "data.txt") {
  const onFinish = vi.fn<(choices: TextImport) => void>()
  const onCancel = vi.fn()
  render(<TextImportWizard bytes={bytes(text)} fileName={fileName} onFinish={onFinish} onCancel={onCancel} />)
  return { onFinish, onCancel }
}

const next = () => fireEvent.click(screen.getByRole("button", { name: "Next >" }))

describe("Text Import Wizard", () => {
  it("steps through Excel's three steps and hands back the choices", () => {
    const { onFinish } = open("id;name;when\n007;Ann;14.03.2026\n")
    expect(screen.getByText("Text Import Wizard - Step 1 of 3")).toBeInTheDocument()
    expect(screen.getByRole("radio", { name: /Delimited/ })).toBeChecked()
    expect(screen.getByRole("button", { name: "< Back" })).toBeDisabled()
    next()
    expect(screen.getByText("Text Import Wizard - Step 2 of 3")).toBeInTheDocument()
    // The semicolon was found; the preview splits on it.
    expect(screen.getByRole("checkbox", { name: "Semicolon" })).toBeChecked()
    expect(screen.getByRole("cell", { name: "Ann" })).toBeInTheDocument()
    next()
    // Column 1 as Text, column 3 as a DMY date.
    fireEvent.click(screen.getByRole("cell", { name: "007" }))
    fireEvent.click(screen.getByRole("radio", { name: "Text" }))
    fireEvent.click(screen.getByRole("cell", { name: "14.03.2026" }))
    fireEvent.click(screen.getByRole("radio", { name: /Date/ }))
    fireEvent.change(screen.getByRole("combobox", { name: "Date order" }), { target: { value: "DMY" } })
    expect(screen.getByRole("columnheader", { name: "DMY" })).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Finish" }))
    expect(onFinish).toHaveBeenCalledWith(
      expect.objectContaining({ kind: "delimited", delimiters: [";"], columns: ["text", "general", "DMY"] }),
    )
  })

  it("spots fixed-width text and lets a break be added and removed", () => {
    const { onFinish } = open("Name      Qty\nApples      3\nPears      12\n")
    expect(screen.getByRole("radio", { name: /Fixed width/ })).toBeChecked()
    next()
    expect(screen.getByRole("separator", { name: "Column break at 10" })).toBeInTheDocument()
    fireEvent.doubleClick(screen.getByRole("separator", { name: "Column break at 10" }))
    expect(screen.queryByRole("separator")).toBeNull()
    fireEvent.click(screen.getByRole("button", { name: "Finish" }))
    expect(onFinish).toHaveBeenCalledWith(expect.objectContaining({ kind: "fixed", breaks: [] }))
  })

  it("start row, other delimiter and Cancel", () => {
    const { onFinish, onCancel } = open("title line\na|b\n")
    fireEvent.change(screen.getByRole("spinbutton"), { target: { value: "2" } })
    next()
    fireEvent.change(screen.getByRole("textbox", { name: "Other delimiter" }), { target: { value: "|" } })
    expect(screen.getByRole("cell", { name: "b" })).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Finish" }))
    expect(onFinish.mock.calls[0][0]).toMatchObject({ startRow: 2 })
    expect(onFinish.mock.calls[0][0].delimiters).toContain("|")
    fireEvent.click(screen.getByRole("button", { name: "Cancel" }))
    expect(onCancel).toHaveBeenCalled()
  })
})
