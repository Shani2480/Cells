import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { searchCommands } from "@/ribbon/tell-me"
import { SheetController } from "@/sheet/controller"
import { QuickAccess, TellMe } from "./QuickAccess"

function envOf(c: SheetController): RibbonEnv {
  return {
    c,
    copy: () => undefined,
    paste: () => undefined,
    dialog: vi.fn(),
    toggleRibbon: () => undefined,
    backstage: () => undefined,
  }
}

describe("Quick Access Toolbar", () => {
  it("shows Undo and Redo, disabled with nothing to undo, and the Customize menu ticks commands on", () => {
    const c = new SheetController()
    const registry = buildRegistry()
    const onChange = vi.fn()
    render(
      <QuickAccess
        registry={registry}
        env={envOf(c)}
        ids={["Undo", "Redo"]}
        onChange={onChange}
        onDone={() => undefined}
      />,
    )
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled()
    fireEvent.click(screen.getByRole("button", { name: "Customize Quick Access Toolbar" }))
    fireEvent.click(screen.getByRole("menuitem", { name: "Save" }))
    expect(onChange).toHaveBeenCalledWith(["Undo", "Redo", "FileSave"])
  })

  it("removes a button from its right-click menu", () => {
    const c = new SheetController()
    const onChange = vi.fn()
    render(
      <QuickAccess
        registry={buildRegistry()}
        env={envOf(c)}
        ids={["Undo", "Redo"]}
        onChange={onChange}
        onDone={() => undefined}
      />,
    )
    fireEvent.contextMenu(screen.getByRole("button", { name: "Redo" }))
    fireEvent.click(screen.getByRole("menuitem", { name: "Remove from Quick Access Toolbar" }))
    expect(onChange).toHaveBeenCalledWith(["Undo"])
  })
})

describe("Search (Tell Me)", () => {
  it("finds commands by name and description, built ones first", () => {
    const c = new SheetController()
    const registry = buildRegistry()
    const names = searchCommands(registry, envOf(c), "freeze").map((cmd) => cmd.label)
    expect(names).toEqual(expect.arrayContaining(["Freeze Top Row", "Freeze First Column"]))
    const merge = searchCommands(registry, envOf(c), "merge center")
    expect(merge[0].label).toMatch(/Merge & Center/)
  })

  it("Enter runs the highlighted result; a menu-only command opens on its tab", () => {
    const c = new SheetController()
    const registry = buildRegistry()
    const onShow = vi.fn(() => true)
    render(<TellMe registry={registry} env={envOf(c)} onShow={onShow} onDone={() => undefined} />)
    const box = screen.getByRole("searchbox", { name: "Search" })
    fireEvent.focus(box)
    fireEvent.change(box, { target: { value: "freeze top row" } })
    fireEvent.keyDown(box, { key: "Enter" })
    expect(c.freeze?.row).toBe(1)
    fireEvent.focus(box)
    fireEvent.change(box, { target: { value: "conditional formatting" } })
    fireEvent.click(screen.getByRole("option", { name: /^Conditional Formatting$/ }))
    expect(onShow).toHaveBeenCalledWith("ConditionalFormattingMenu")
  })
})
