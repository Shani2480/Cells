import { act, fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { MoveCopyDialog, UnhideSheetDialog } from "./SheetDialogs"
import { SheetTabs } from "./SheetTabs"

function setup() {
  const c = new SheetController()
  const dialogs: string[] = []
  const alerts: string[] = []
  c.shell = { ...c.shell, alert: (m) => alerts.push(m) }
  render(<SheetTabs controller={c} onDialog={(name) => dialogs.push(name)} onFocusGrid={() => undefined} />)
  const tab = (name: string) => screen.getByRole("tab", { name })
  return { c, dialogs, alerts, tab }
}

describe("sheet tab bar", () => {
  it("New sheet adds a tab after the active one and shows it", () => {
    const { c, tab } = setup()
    fireEvent.click(screen.getByRole("button", { name: "New sheet" }))
    expect(tab("Sheet2").getAttribute("aria-selected")).toBe("true")
    expect(c.sheet.name).toBe("Sheet2")
    fireEvent.pointerDown(tab("Sheet1"), { button: 0, pointerId: 1 })
    fireEvent.pointerUp(tab("Sheet1"), { pointerId: 1 })
    expect(c.sheet.name).toBe("Sheet1")
  })

  it("Shift+click groups tabs; the tab menu then offers Ungroup Sheets", () => {
    const { c, tab } = setup()
    act(() => {
      c.sheets.insert(1)
      c.sheets.insert(2)
      c.activate(0)
    })
    fireEvent.pointerDown(tab("Sheet3"), { button: 0, pointerId: 1, shiftKey: true })
    expect(c.group.indexes()).toEqual([0, 1, 2])
    expect(tab("Sheet2").hasAttribute("data-grouped")).toBe(true)
    fireEvent.contextMenu(tab("Sheet2"), { clientX: 40, clientY: 500 })
    fireEvent.click(screen.getByRole("menuitem", { name: "Ungroup Sheets" }))
    expect(c.group.active).toBe(false)
  })

  it("double-click renames in place: Enter keeps the name, Esc puts the old one back, bad names are refused", () => {
    const { c, tab, alerts } = setup()
    act(() => c.sheets.insert(1))
    fireEvent.doubleClick(tab("Sheet1"))
    const box = screen.getByRole("textbox", { name: "Sheet name" })
    fireEvent.change(box, { target: { value: "Budget" } })
    fireEvent.keyDown(box, { key: "Enter" })
    expect(c.wb.sheets[0].name).toBe("Budget")
    fireEvent.doubleClick(tab("Budget"))
    const again = screen.getByRole("textbox", { name: "Sheet name" })
    fireEvent.change(again, { target: { value: "sheet2" } })
    fireEvent.keyDown(again, { key: "Enter" })
    expect(alerts[0]).toMatch(/already taken/)
    fireEvent.keyDown(again, { key: "Escape" })
    expect(screen.queryByRole("textbox", { name: "Sheet name" })).toBeNull()
    expect(c.wb.sheets[0].name).toBe("Budget")
  })

  it("the tab menu has Excel's items; what isn't built is disabled", () => {
    const { dialogs, tab } = setup()
    fireEvent.contextMenu(tab("Sheet1"), { clientX: 40, clientY: 500 })
    for (const label of [
      "Insert...",
      "Delete",
      "Rename",
      "Move or Copy...",
      "Protect Sheet...",
      "Tab Color",
      "Hide",
      "Select All Sheets",
    ])
      expect((screen.getByRole("menuitem", { name: label }) as HTMLButtonElement).disabled).toBe(false)
    for (const label of ["View Code", "Unhide..."])
      expect((screen.getByRole("menuitem", { name: label }) as HTMLButtonElement).disabled).toBe(true)
    fireEvent.click(screen.getByRole("menuitem", { name: "Move or Copy..." }))
    expect(dialogs).toEqual(["moveCopySheet"])
  })

  it("Tab Color colours the tab", () => {
    const { c, tab } = setup()
    fireEvent.contextMenu(tab("Sheet1"), { clientX: 40, clientY: 500 })
    fireEvent.click(screen.getByRole("menuitem", { name: "Tab Color" }))
    fireEvent.click(screen.getAllByRole("menuitemradio", { name: "Red" })[0])
    expect(c.sheet.tabColor).toBe("#FF0000")
    expect(tab("Sheet1").getAttribute("data-colored")).toBe("dark")
  })

  it("hidden sheets have no tab", () => {
    const { c, tab } = setup()
    act(() => c.sheets.insert(1))
    act(() => c.sheets.hide(1))
    expect(screen.queryByRole("tab", { name: "Sheet2" })).toBeNull()
    expect(tab("Sheet1")).toBeTruthy()
  })
})

describe("sheet dialogs", () => {
  it("Move or Copy: before a sheet, or a copy", () => {
    const c = new SheetController()
    c.sheets.insert(1)
    c.sheets.insert(2)
    c.activate(0)
    const { unmount } = render(<MoveCopyDialog controller={c} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("option", { name: "(move to end)" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet2", "Sheet3", "Sheet1"])
    unmount()
    render(<MoveCopyDialog controller={c} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("option", { name: "Sheet2" }))
    fireEvent.click(screen.getByRole("checkbox"))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1 (2)", "Sheet2", "Sheet3", "Sheet1"])
  })

  it("Unhide lists the hidden sheets", () => {
    const c = new SheetController()
    c.sheets.insert(1)
    c.sheets.hide(0)
    render(<UnhideSheetDialog controller={c} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.wb.sheets[0].hidden).toBe(false)
    expect(c.sheet.name).toBe("Sheet1")
  })
})
