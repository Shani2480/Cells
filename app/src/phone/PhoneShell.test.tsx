import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { DATA, HOME, TABS, type Tab } from "@/ribbon/tabs"
import { SheetController } from "@/sheet/controller"
import { actionSheetItems } from "./ActionSheet"
import { dismisses } from "./BottomSheet"
import { paletteGroups } from "./CommandSheet"
import { FORMULA_KEYS, PhoneShell, typeAtCaret } from "./PhoneShell"

function setup(tabs: Tab[] = TABS) {
  const c = new SheetController()
  const registry = buildRegistry()
  const dialogs: string[] = []
  const env: RibbonEnv = {
    c,
    copy: () => undefined,
    paste: () => undefined,
    dialog: (name) => dialogs.push(name),
    toggleRibbon: () => undefined,
    backstage: () => undefined,
  }
  let file = 0
  render(
    <PhoneShell
      controller={c}
      env={env}
      registry={registry}
      tabs={tabs}
      title="Q1 Report"
      formulaBar={null}
      onFile={() => file++}
      onFocusGrid={() => undefined}
    >
      <div data-testid="grid" />
    </PhoneShell>,
  )
  return { c, registry, dialogs, files: () => file }
}

const palette = () => screen.getByRole("dialog", { name: "Commands" })

describe("the phone shell", () => {
  it("shows the workbook's name, Undo and Redo, the sheet switcher and the tools", () => {
    const { c } = setup()
    expect(screen.getByText("Q1 Report")).toBeInTheDocument()
    expect(screen.getByRole("button", { name: "Undo" })).toBeDisabled()
    for (const name of ["Home", "Insert", "Formulas", "Data", "View"])
      expect(screen.getByRole("button", { name })).toBeInTheDocument()
    expect(screen.getByRole("button", { name: "Sheet: Sheet1" })).toBeInTheDocument()
    // Finger-sized headings while it is shown.
    expect(c.largeHeaders).toBe(true)
  })

  it("lists a tab's built commands by group, leaving out what isn't built", () => {
    const registry = buildRegistry()
    const groups = paletteGroups(HOME, registry)
    const ids = groups.flatMap((g) => g.rows.map((r) => r.id))
    expect(groups.map((g) => g.label)).toContain("Font")
    expect(ids).toContain("Bold")
    expect(ids.every((id) => registry.isBuilt(id))).toBe(true)
    expect(groups.every((g) => g.rows.length > 0)).toBe(true)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it("a tool opens the palette on its tab; a command runs and shows it's on", () => {
    const { c } = setup()
    fireEvent.click(screen.getByRole("button", { name: "Home" }))
    const bold = within(palette()).getByRole("button", { name: "Bold" })
    fireEvent.click(bold)
    expect(c.sheet.cells.get(0, 0)?.style).toBeDefined()
    expect(within(palette()).getByRole("button", { name: "Bold" })).toHaveAttribute("aria-pressed", "true")
    expect(screen.getByRole("button", { name: "Undo" })).not.toBeDisabled()
  })

  it("the tab picker switches the palette's tab", () => {
    setup()
    fireEvent.click(screen.getByRole("button", { name: "Home" }))
    fireEvent.click(within(palette()).getByRole("button", { name: "Home" }))
    fireEvent.click(within(palette()).getByRole("option", { name: "Data" }))
    expect(within(palette()).getByRole("button", { name: "Data" })).toBeInTheDocument()
    expect(within(palette()).getByRole("region", { name: DATA.groups[0].label })).toBeInTheDocument()
  })

  it("a dropdown opens its menu in place, with a Back row", () => {
    setup()
    fireEvent.click(screen.getByRole("button", { name: "Home" }))
    const groups = paletteGroups(HOME, buildRegistry())
    const borders = groups.flatMap((g) => g.rows).find((r) => r.id === "BordersGallery")
    expect(borders).toBeDefined()
    fireEvent.click(within(palette()).getByRole("button", { name: "Borders options" }))
    expect(within(palette()).getByRole("menu")).toBeInTheDocument()
    fireEvent.click(within(palette()).getByRole("button", { name: /Borders|Bottom Border/ }))
    expect(within(palette()).queryByRole("menu")).toBeNull()
  })

  it("the sheet switcher lists the sheets, switches, and adds one", async () => {
    const { c } = setup()
    act(() => c.sheets.insert(1))
    act(() => c.activate(0))
    fireEvent.click(screen.getByRole("button", { name: "Sheet: Sheet1" }))
    const list = screen.getByRole("dialog", { name: "Sheets" })
    expect(within(list).getByRole("option", { name: "Sheet1" })).toHaveAttribute("aria-selected", "true")
    fireEvent.click(within(list).getByRole("option", { name: "Sheet2" }))
    expect(c.sheet.name).toBe("Sheet2")
    await waitFor(() => expect(screen.queryByRole("dialog", { name: "Sheets" })).toBeNull())
    fireEvent.click(screen.getByRole("button", { name: "Sheet: Sheet2" }))
    fireEvent.click(within(screen.getByRole("dialog", { name: "Sheets" })).getByRole("button", { name: "New Sheet" }))
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "Sheet2", "Sheet3"])
    expect(c.sheet.name).toBe("Sheet3")
  })

  it("while editing, the formula keys replace the bottom bar and type at the caret", () => {
    const { c, dialogs } = setup()
    act(() => c.beginEdit("enter", "=SUM"))
    expect(screen.queryByRole("button", { name: "Home" })).toBeNull()
    const keys = screen.getByRole("toolbar", { name: "Formula keys" })
    for (const key of FORMULA_KEYS) expect(within(keys).getByRole("button", { name: key })).toBeInTheDocument()
    fireEvent.click(within(keys).getByRole("button", { name: "(" }))
    expect(c.edit?.text).toBe("=SUM(")
    act(() => c.setEditText("=SUM(A1)", 6, 8))
    act(() => typeAtCaret(c, ":"))
    expect(c.edit?.text).toBe("=SUM(A:")
    fireEvent.click(within(keys).getByRole("button", { name: "Insert Function" }))
    expect(dialogs).toEqual(["insertFunction"])
    act(() => c.setEditText("=1+2"))
    fireEvent.click(within(keys).getByRole("button", { name: "Enter" }))
    expect(c.edit).toBeNull()
    expect(c.sheet.cells.get(0, 0)?.value).toBe(3)
    expect(screen.getByRole("button", { name: "Home" })).toBeInTheDocument()
  })

  it("the File button opens the File view", () => {
    const { files } = setup()
    fireEvent.click(screen.getByRole("button", { name: "File" }))
    expect(files()).toBe(1)
  })
})

describe("bottom sheets", () => {
  it("close when dragged down far enough, or flicked", () => {
    expect(dismisses(40, 0.1)).toBe(false)
    expect(dismisses(120, 0)).toBe(true)
    expect(dismisses(30, 0.9)).toBe(true)
    expect(dismisses(8, 2)).toBe(false) // a tap that jittered
    expect(dismisses(-50, 0)).toBe(false)
  })

  it("drag down from the tab picker too, without pressing it", async () => {
    setup()
    fireEvent.click(screen.getByRole("button", { name: "Home" }))
    const picker = within(palette()).getByRole("button", { name: "Home" })
    fireEvent.pointerDown(picker, { clientY: 100, pointerId: 1 })
    fireEvent.pointerMove(picker, { clientY: 160, pointerId: 1 })
    fireEvent.pointerMove(picker, { clientY: 260, pointerId: 1 })
    fireEvent.pointerUp(picker, { clientY: 260, pointerId: 1 })
    fireEvent.click(picker)
    // The tab list didn't open, and the sheet goes.
    expect(within(palette()).queryByRole("listbox", { name: "Tabs" })).toBeNull()
    await waitFor(() => expect(screen.queryByRole("dialog", { name: "Commands" })).toBeNull())
  })

  it("a tap on the tab picker still opens the tab list", () => {
    setup()
    fireEvent.click(screen.getByRole("button", { name: "Home" }))
    const picker = within(palette()).getByRole("button", { name: "Home" })
    fireEvent.pointerDown(picker, { clientY: 100, pointerId: 1 })
    fireEvent.pointerMove(picker, { clientY: 103, pointerId: 1 })
    fireEvent.pointerUp(picker, { clientY: 103, pointerId: 1 })
    fireEvent.click(picker)
    expect(within(palette()).getByRole("listbox", { name: "Tabs" })).toBeInTheDocument()
  })

  it("close on Esc (Android's Back)", async () => {
    setup()
    fireEvent.click(screen.getByRole("button", { name: "Home" }))
    fireEvent.keyDown(window, { key: "Escape" })
    await waitFor(() => expect(screen.queryByRole("dialog", { name: "Commands" })).toBeNull())
  })
})

describe("the long-press sheet", () => {
  const item = (label: string, extra: object = {}) => ({ kind: "item" as const, label, onSelect: () => {}, ...extra })
  it("puts Cut, Copy and Paste in the quick row and drops unbuilt items and empty groups", () => {
    const sep = { kind: "separator" as const }
    const { quick, groups } = actionSheetItems([
      item("Cut", { icon: "✂" }),
      item("Copy", { icon: "⧉" }),
      item("Paste", { icon: "📋" }),
      item("Paste Special...", { disabled: true }),
      sep,
      item("Insert..."),
      sep,
      item("Quick Analysis", { disabled: true }),
      sep,
      sep,
      item("Format Cells..."),
    ])
    expect(quick.map((i) => i.label)).toEqual(["Cut", "Copy", "Paste"])
    expect(groups.map((g) => g.map((i) => i.label))).toEqual([["Insert..."], ["Format Cells..."]])
  })
})
