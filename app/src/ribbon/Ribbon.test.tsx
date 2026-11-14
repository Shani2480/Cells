import { parseRange, selectRange, writeInputOp } from "@cells/core"
import { act, fireEvent, render, screen, within } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { TooltipProvider } from "@/ui/tooltip"
import { buildRegistry } from "./commands"
import type { MenuEntry, RibbonEnv } from "./registry"
import { Ribbon } from "./Ribbon"
import { TABLE_DESIGN, TABS, controlIds } from "./tabs"

function setup() {
  const c = new SheetController()
  const dialogs: string[] = []
  const env: RibbonEnv = {
    c,
    copy: () => undefined,
    paste: () => undefined,
    dialog: (name) => dialogs.push(name),
    toggleRibbon: () => undefined,
    backstage: () => undefined,
  }
  const registry = buildRegistry()
  render(
    <TooltipProvider>
      <Ribbon env={env} registry={registry} tabs={TABS} onFile={() => dialogs.push("file")} />
    </TooltipProvider>,
  )
  const style = (ref: string) => {
    const r = parseRange(ref)!
    return c.wb.styles.get(c.sheet.styleAt(r.r0, r.c0))
  }
  const panel = () => within(screen.getByRole("tabpanel"))
  /** The visible ribbon's control (the measuring copy is hidden). */
  const control = (id: string) => document.querySelector<HTMLElement>(`.ribbon-groups [data-command="${id}"]`)!
  return { c, env, registry, dialogs, style, panel, control }
}

describe("command registry", () => {
  it("every control on every tab and in every menu has a command with a name", () => {
    const registry = buildRegistry()
    const env = { c: new SheetController() } as RibbonEnv
    const missing: string[] = []
    const check = (id: string) => {
      const command = registry.get(id)
      // A label equal to a camel-case id means no name was found.
      if (!command || !command.label || (command.label === id && /[a-z][A-Z]/.test(id))) missing.push(id)
      const menu = command?.menu ? (typeof command.menu === "function" ? command.menu(env) : command.menu) : []
      for (const entry of menu as MenuEntry[]) if (typeof entry === "string" && entry !== "-") check(entry)
    }
    for (const tab of [...TABS, TABLE_DESIGN]) for (const id of controlIds(tab)) check(id)
    expect([...new Set(missing)]).toEqual([])
  })

  it("leaves out Excel's cloud-only groups", () => {
    const groups = TABS.flatMap((t) => t.groups.map((g) => g.id))
    for (const cloud of [
      "GroupClassifyLabelProtect",
      "GroupAIAssistance",
      "GroupIdeas",
      "GroupLinkedEntityConvert",
      "GroupInsights",
      "GroupLanguage",
    ])
      expect(groups).not.toContain(cloud)
  })
})

describe("ribbon", () => {
  it("Bold on the Home tab formats the selection and shows as pressed", () => {
    const { c, style, control } = setup()
    act(() => c.setSelection(selectRange(parseRange("B2:C3")!)))
    const bold = control("Bold")
    expect(bold.getAttribute("aria-pressed")).toBe("false")
    fireEvent.click(bold)
    expect(style("C3").bold).toBe(true)
    expect(control("Bold").getAttribute("aria-pressed")).toBe("true")
  })

  it("a command that is not built yet is disabled and does nothing", () => {
    const { control } = setup()
    fireEvent.click(screen.getByRole("tab", { name: "Insert" }))
    const pivot = control("PivotTableInsert")
    expect(pivot.getAttribute("aria-disabled")).toBe("true")
    fireEvent.click(pivot)
    expect(document.querySelector(".popover")).toBeNull()
  })

  it("dialog launchers and menu items open the shell's dialogs", () => {
    const { dialogs, panel } = setup()
    fireEvent.click(panel().getByRole("button", { name: "Font Settings" }))
    expect(dialogs).toContain("formatCells:font")
  })

  it("the Font Size box applies a typed size on Enter", () => {
    const { style } = setup()
    const box = screen.getAllByRole("textbox", { name: "Font Size" })[0]
    fireEvent.focus(box)
    fireEvent.change(box, { target: { value: "14" } })
    fireEvent.keyDown(box, { key: "Enter" })
    expect(style("A1").fontSize).toBe(14)
  })

  it("KeyTips: Alt, H, 1 makes the selection bold; Alt, H, A, C centers it", () => {
    const { style } = setup()
    const press = (key: string) => {
      fireEvent.keyDown(window, { key })
      if (key === "Alt") fireEvent.keyUp(window, { key })
    }
    act(() => press("Alt"))
    expect(screen.getAllByText("H").some((el) => el.classList.contains("keytip"))).toBe(true)
    act(() => press("h"))
    act(() => press("1"))
    expect(style("A1").bold).toBe(true)
    act(() => press("Alt"))
    act(() => press("h"))
    act(() => press("a"))
    act(() => press("c"))
    expect(style("A1").hAlign).toBe("center")
  })

  it("Merge & Center on the ribbon merges, and Merge Across from its menu", () => {
    const { c, control } = setup()
    c.wb.transact([writeInputOp(0, c.sheet, c.wb.styles, 0, 0, "x")], { label: "t" })
    act(() => c.setSelection(selectRange(parseRange("A1:C2")!)))
    fireEvent.click(control("MergeCenterMenu").querySelector(".ribbon-small-main")!)
    expect(c.sheet.merges.all).toEqual([parseRange("A1:C2")])
  })

  it("View › Gridlines and Headings toggle the sheet's display", () => {
    const { c } = setup()
    fireEvent.click(screen.getByRole("tab", { name: "View" }))
    const grid = document.querySelector<HTMLElement>('.ribbon-groups [data-command="GridlinesExcel"]')!
    fireEvent.click(grid)
    expect(c.showGridlines).toBe(false)
    fireEvent.click(document.querySelector<HTMLElement>('.ribbon-groups [data-command="ViewHeadings"]')!)
    expect(c.showHeadings).toBe(false)
  })

  it("double-clicking a tab collapses the ribbon to its tabs", () => {
    setup()
    fireEvent.doubleClick(screen.getByRole("tab", { name: "Home" }))
    expect(screen.queryByRole("tabpanel")).toBeNull()
    fireEvent.click(screen.getByRole("tab", { name: "Insert" }))
    expect(screen.getByRole("tabpanel", { name: "Insert" })).toBeTruthy()
    localStorage.clear()
  })
})
