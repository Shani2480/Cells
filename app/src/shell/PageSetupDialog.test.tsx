import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { BAD_REFERENCE, PageSetupDialog } from "./PageSetupDialog"
import { codesToNames, namesToCodes } from "@/sheet/header-footer"

describe("Page Setup dialog", () => {
  it("applies every tab's changes on OK as one undo step", () => {
    const c = new SheetController()
    c.print.defaultPaper = 9
    const onClose = vi.fn()
    render(<PageSetupDialog controller={c} fileName="Book1.xlsx" onClose={onClose} />)
    fireEvent.click(screen.getByRole("radio", { name: "Landscape" }))
    fireEvent.click(screen.getByRole("radio", { name: "Fit to:" }))
    fireEvent.click(screen.getByRole("tab", { name: "Sheet" }))
    fireEvent.change(screen.getByLabelText("Print area:"), { target: { value: "$A$1:$H$60" } })
    fireEvent.change(screen.getByLabelText("Rows to repeat at top:"), { target: { value: "$1:$2" } })
    fireEvent.change(screen.getByLabelText("Columns to repeat at left:"), { target: { value: "$A:$A" } })
    fireEvent.click(screen.getByRole("checkbox", { name: "Gridlines" }))
    fireEvent.click(screen.getByRole("radio", { name: "Over, then down" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    const s = c.print.setup()
    expect(s).toMatchObject({
      orientation: "landscape",
      fitTo: { width: 1, height: 1 },
      printArea: [{ r0: 0, c0: 0, r1: 59, c1: 7 }],
      titleRows: [0, 1],
      titleCols: [0, 0],
      gridlines: true,
      order: "overThenDown",
    })
    c.undo()
    expect(c.print.setup()).toMatchObject({ orientation: "portrait", printArea: null, gridlines: false })
  })

  it("refuses a print title it can't read and goes back to the Sheet tab", () => {
    const c = new SheetController()
    c.shell.alert = vi.fn()
    const onClose = vi.fn()
    render(<PageSetupDialog controller={c} tab="sheet" fileName="Book1.xlsx" onClose={onClose} />)
    fireEvent.change(screen.getByLabelText("Rows to repeat at top:"), { target: { value: "top rows" } })
    fireEvent.click(screen.getByRole("tab", { name: "Margins" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.shell.alert).toHaveBeenCalledWith(BAD_REFERENCE)
    expect(onClose).not.toHaveBeenCalled()
    expect(screen.getByRole("tab", { name: "Sheet" })).toHaveAttribute("aria-selected", "true")
  })

  it("shows header codes by name and keeps them as letters", () => {
    expect(codesToNames("Page &P of &N, &A")).toBe("Page &[Page] of &[Pages], &[Tab]")
    expect(namesToCodes("Page &[Page] of &[Pages], &[Tab]")).toBe("Page &P of &N, &A")
    expect(codesToNames("A && B")).toBe("A && B")
  })
})
