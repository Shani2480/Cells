import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { SHEET_MESSAGES } from "@/sheet/sheets"
import { ProtectSheetDialog, ProtectWorkbookDialog } from "./ProtectDialogs"

function setup() {
  const c = new SheetController()
  const messages: string[] = []
  c.shell = { ...c.shell, alert: (m) => messages.push(m) }
  let closed = 0
  render(<ProtectSheetDialog controller={c} onClose={() => closed++} />)
  return { c, messages, closed: () => closed }
}

describe("Protect Sheet dialog", () => {
  it("lists Excel's options with only selecting allowed by default", () => {
    setup()
    expect(screen.getByLabelText("Select locked cells")).toBeChecked()
    expect(screen.getByLabelText("Select unlocked cells")).toBeChecked()
    expect(screen.getByLabelText("Format cells")).not.toBeChecked()
    expect(screen.getByLabelText("Edit scenarios")).not.toBeChecked()
    expect(screen.getByLabelText(/Password to unprotect sheet/)).toHaveFocus()
  })

  it("protects with the chosen options and no password", () => {
    const { c, closed } = setup()
    fireEvent.click(screen.getByLabelText("Insert rows"))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.protection).toMatchObject({ insertRows: true, formatCells: false })
    expect(c.sheet.protection?.password).toBeUndefined()
    expect(closed()).toBe(1)
  })

  it("asks to confirm a password, and a mismatch shows Excel's message", () => {
    const { c, messages } = setup()
    fireEvent.change(screen.getByLabelText(/Password to unprotect sheet/), { target: { value: "abc" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    fireEvent.change(screen.getByLabelText(/Reenter password/), { target: { value: "abd" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(messages).toEqual([SHEET_MESSAGES.passwordMismatch])
    expect(c.sheet.protection).toBeNull()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    fireEvent.change(screen.getByLabelText(/Reenter password/), { target: { value: "abc" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.protection?.password).toBeDefined()
  })
})

describe("Protect Workbook dialog", () => {
  it("protects the structure with a confirmed password; Windows is greyed out", () => {
    const c = new SheetController()
    let closed = 0
    render(<ProtectWorkbookDialog controller={c} onClose={() => closed++} />)
    expect(screen.getByLabelText("Structure")).toBeChecked()
    expect(screen.getByLabelText("Windows")).toBeDisabled()
    fireEvent.change(screen.getByLabelText(/Password \(optional\)/), { target: { value: "pw" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    fireEvent.change(screen.getByLabelText(/Reenter password/), { target: { value: "pw" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheets.isWorkbookProtected()).toBe(true)
    expect(c.sheets.workbookNeedsPassword()).toBe(true)
    expect(closed).toBe(1)
  })
})
