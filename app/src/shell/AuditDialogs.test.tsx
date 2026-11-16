import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { SheetController } from "@/sheet/controller"
import { CreateNamesDialog, ErrorCheckingDialog, EvaluateFormulaDialog, NameManagerDialog } from "./AuditDialogs"
import { WatchWindow } from "./WatchWindow"

function setup(cells: [number, number, string][]) {
  const c = new SheetController()
  for (const [row, col, text] of cells) {
    c.goTo({ r0: row, c0: col, r1: row, c1: col })
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  c.wb.calculate(true)
  return c
}

describe("Name Manager", () => {
  it("lists each name's value, refers to and scope; edits Refers To in place; deletes", () => {
    const c = setup([
      [0, 0, "Rent"],
      [1, 0, "Food"],
      [0, 1, "1200"],
    ])
    c.names.define("Items", "=Sheet1!$A$1:$A$2")
    c.names.define("Cost", "=Sheet1!$B$1")
    render(<NameManagerDialog controller={c} onClose={() => undefined} />)
    expect(screen.getByText('{"Rent";"Food"}')).toBeInTheDocument()
    expect(screen.getByText("=Sheet1!$B$1")).toBeInTheDocument()
    fireEvent.click(screen.getByText("Items"))
    const field = screen.getByLabelText("Refers to:")
    fireEvent.change(field, { target: { value: "=Sheet1!$A$1" } })
    fireEvent.keyDown(field, { key: "Enter" })
    expect(c.names.all().find((n) => n.name === "Items")?.ref).toBe("Sheet1!$A$1")
    fireEvent.click(screen.getByRole("button", { name: "Delete" }))
    expect(c.names.all().map((n) => n.name)).toEqual(["Cost"])
    fireEvent.click(screen.getByText("Cost"))
    fireEvent.click(screen.getByRole("button", { name: "Edit..." }))
    expect(screen.getByText("Edit Name")).toBeInTheDocument()
  })

  it("Ctrl+F3 asks for it, Ctrl+Shift+F3 for Create from Selection", () => {
    const c = new SheetController()
    const openDialog = vi.fn()
    c.shell = { ...c.shell, openDialog }
    c.keyReady({ key: "F3", primary: true, shift: false, alt: false })
    c.keyReady({ key: "F3", primary: true, shift: true, alt: false })
    expect(openDialog.mock.calls).toEqual([["nameManager"], ["createNames"]])
  })
})

describe("Create Names from Selection", () => {
  it("starts from Excel's guess and makes the names", () => {
    const c = setup([
      [0, 0, "Qty"],
      [1, 0, "3"],
    ])
    c.goTo({ r0: 0, c0: 0, r1: 1, c1: 0 })
    const close = vi.fn()
    render(<CreateNamesDialog controller={c} onClose={close} />)
    expect(screen.getByLabelText("Top row")).toBeChecked()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.names.all()).toEqual([{ name: "Qty", ref: "Sheet1!$A$2" }])
    expect(close).toHaveBeenCalled()
  })
})

describe("Evaluate Formula", () => {
  it("underlines the next part and evaluates it, step by step, then restarts", () => {
    const c = setup([
      [0, 0, "2"],
      [0, 1, "=A1*3+1"],
    ])
    c.goTo({ r0: 0, c0: 1, r1: 0, c1: 1 })
    render(<EvaluateFormulaDialog controller={c} onClose={() => undefined} />)
    const box = () => screen.getByRole("textbox", { name: "Evaluation" })
    const next = () => box().querySelector("u")?.textContent ?? null
    expect(screen.getByText("Sheet1!$B$1")).toBeInTheDocument()
    expect([box().textContent, next()]).toEqual(["=A1*3+1", "A1"])
    const evaluate = () => fireEvent.click(screen.getByRole("button", { name: /Evaluate|Restart/ }))
    evaluate()
    expect([box().textContent, next()]).toEqual(["=2*3+1", "2*3"])
    evaluate()
    evaluate()
    expect([box().textContent, next()]).toEqual(["=7", null])
    expect(screen.getByRole("button", { name: "Restart" })).toBeInTheDocument()
    evaluate()
    expect(box().textContent).toBe("=A1*3+1")
  })
})

describe("Error Checking", () => {
  it("explains the error, goes to the next problem and offers Convert to Number", () => {
    const c = setup([
      [0, 0, "0"],
      [1, 1, "=1/A1"],
      [2, 2, "'15"],
    ])
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    render(<ErrorCheckingDialog controller={c} onClose={() => undefined} onEvaluate={() => undefined} />)
    expect(screen.getByText("Error in cell B2")).toBeInTheDocument()
    expect(screen.getByText("Divide by Zero Error")).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Next" }))
    expect(screen.getByText("Number Stored as Text")).toBeInTheDocument()
    expect(c.selection.active).toEqual({ row: 2, col: 2 })
    fireEvent.click(screen.getByRole("button", { name: "Convert to Number" }))
    expect(c.sheet.cells.value(2, 2)).toBe(15)
  })
})

describe("the auditing commands and the Watch Window", () => {
  it("Trace Precedents draws arrows, Watch Window opens and shows the watched cells' values", () => {
    const c = setup([
      [0, 0, "4"],
      [0, 1, "=A1*2"],
    ])
    const registry = buildRegistry()
    const env = { c, dialog: vi.fn() } as unknown as RibbonEnv
    c.goTo({ r0: 0, c0: 1, r1: 0, c1: 1 })
    registry.get("TracePrecedents")!.run!(env)
    expect(c.audit.arrows()).toHaveLength(1)
    registry.get("TraceRemoveArrowsMenu")!.run!(env)
    expect(c.audit.arrows()).toHaveLength(0)
    registry.get("FormulaEvaluate")!.run!(env)
    expect(env.dialog).toHaveBeenCalledWith("evaluateFormula")
    registry.get("WatchWindow")!.run!(env)
    expect(registry.get("WatchWindow")!.pressed!(env)).toBe(true)
    render(<WatchWindow controller={c} onDone={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "Add Watch" }))
    expect(screen.getByText("8")).toBeInTheDocument()
    expect(screen.getByText("=A1*2")).toBeInTheDocument()
  })
})
