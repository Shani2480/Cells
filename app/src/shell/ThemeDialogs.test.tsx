import { fireEvent, render, screen } from "@testing-library/react"
import { afterEach, describe, expect, it, vi } from "vitest"
import { OFFICE_THEME } from "@cells/core"
import type { RibbonEnv } from "@/ribbon/registry"
import { themeCommands } from "@/ribbon/theme-commands"
import { useSettings } from "@/settings/settings"
import { SheetController } from "@/sheet/controller"
import { putInList, ThemeColorsDialog, ThemeFontsDialog } from "./ThemeDialogs"

const panelOf = (id: string) => {
  const command = themeCommands().find((x) => x.id === id)!
  const entry = command.menu![0] as { panel: (env: RibbonEnv, close: () => void) => React.ReactNode }
  return entry.panel
}

const envFor = (c: SheetController) =>
  ({ c, dialog: vi.fn(), copy() {}, paste() {}, toggleRibbon() {}, backstage() {} }) as unknown as RibbonEnv

describe("custom theme colours and fonts", () => {
  afterEach(() => useSettings.setState({ themeColors: [], themeFonts: [], savedThemes: [] }))

  it("Create New Theme Colors keeps a named set and gives the workbook its colours", () => {
    const c = new SheetController()
    const onClose = vi.fn()
    render(<ThemeColorsDialog controller={c} onClose={onClose} />)
    expect(screen.getByLabelText("Name:")).toHaveProperty("value", "Custom 1")
    fireEvent.click(screen.getByRole("button", { name: "Accent 1" }))
    fireEvent.click(screen.getByRole("menuitemradio", { name: "Red" }))
    fireEvent.change(screen.getByLabelText("Name:"), { target: { value: "Harbour" } })
    fireEvent.click(screen.getByRole("button", { name: "Save" }))
    expect(onClose).toHaveBeenCalled()
    const set = useSettings.getState().themeColors
    expect(set).toEqual([{ name: "Harbour", colors: { ...OFFICE_THEME.colors, accent1: "FF0000" } }])
    expect(c.wb.theme).toMatchObject({ name: "Office Theme", colorsName: "Harbour" })
    expect(c.wb.theme.colors.accent1).toBe("FF0000")
    c.undo()
    expect(c.wb.theme.colors.accent1).toBe(OFFICE_THEME.colors.accent1)
  })

  it("Create New Theme Fonts keeps a pair; editing one renames it in place", () => {
    const c = new SheetController()
    useSettings.setState({ themeFonts: [{ name: "Custom 1", major: "Georgia", minor: "Verdana" }] })
    render(<ThemeFontsDialog controller={c} edit="Custom 1" onClose={() => {}} />)
    expect(screen.getByText("Edit Theme Fonts")).toBeTruthy()
    expect(screen.getByLabelText("Heading font:")).toHaveProperty("value", "Georgia")
    fireEvent.change(screen.getByLabelText("Body font:"), { target: { value: "Arial" } })
    fireEvent.change(screen.getByLabelText("Name:"), { target: { value: "Letters" } })
    fireEvent.click(screen.getByRole("button", { name: "Save" }))
    expect(useSettings.getState().themeFonts).toEqual([{ name: "Letters", major: "Georgia", minor: "Arial" }])
    expect(c.wb.theme).toMatchObject({ fonts: { major: "Georgia", minor: "Arial" }, fontsName: "Letters" })
  })

  it("the galleries list the person's sets under Custom, above Office's; right-click edits or deletes", () => {
    const c = new SheetController()
    const env = envFor(c)
    const colors = { ...OFFICE_THEME.colors, accent1: "1F6F8B" }
    useSettings.setState({ themeColors: [{ name: "Harbour", colors }] })
    const close = vi.fn()
    render(<>{panelOf("ThemeColorsGallery")(env, close)}</>)
    expect(screen.getAllByText(/^(Custom|Office)$/).map((e) => e.textContent)).toEqual(["Custom", "Office", "Office"])
    fireEvent.contextMenu(screen.getByRole("menuitemradio", { name: "Harbour" }))
    fireEvent.click(screen.getByText("Edit..."))
    expect(env.dialog).toHaveBeenCalledWith("themeColors:Harbour")
    fireEvent.click(screen.getByRole("menuitemradio", { name: "Harbour" }))
    expect(c.wb.theme).toMatchObject({ colorsName: "Harbour" })
    expect(c.wb.theme.colors.accent1).toBe("1F6F8B")
  })

  it("Save Current Theme writes a .thmx named as the file and lists it under Custom", async () => {
    const c = new SheetController()
    let written: Uint8Array | null = null
    const env = {
      ...envFor(c),
      files: {
        saveOther: vi.fn(async (_name: string, _type: unknown, make: (n: string) => Promise<Uint8Array>) => {
          written = await make("Harbour")
          return "Harbour"
        }),
      },
    } as unknown as RibbonEnv
    themeCommands().find((x) => x.id === "ThemeSaveCurrentTheme")!.run!(env)
    await vi.waitFor(() => expect(useSettings.getState().savedThemes).toHaveLength(1))
    expect(env.files!.saveOther).toHaveBeenCalledWith("Theme1.thmx", expect.anything(), expect.any(Function))
    const { themeFromFile } = await import("@cells/io")
    expect(await themeFromFile(written!)).toMatchObject({ name: "Harbour" })
    expect(useSettings.getState().savedThemes[0]).toMatchObject({ name: "Harbour", fonts: OFFICE_THEME.fonts })
    expect(c.wb.theme.name).toBe("Harbour")
    render(<>{panelOf("ThemesGallery")(env, () => {})}</>)
    expect(screen.getByRole("menuitemradio", { name: "Harbour" }).getAttribute("aria-checked")).toBe("true")
  })

  it("a set takes the place of the one it replaces", () => {
    const list = [
      { name: "A", major: "x", minor: "x" },
      { name: "B", major: "x", minor: "x" },
      { name: "C", major: "x", minor: "x" },
    ]
    expect(putInList(list, { name: "B2", major: "y", minor: "y" }, "B").map((x) => x.name)).toEqual(["A", "B2", "C"])
    expect(putInList(list, { name: "C", major: "y", minor: "y" }, "A").map((x) => x.name)).toEqual(["C", "B"])
    expect(putInList(list, { name: "D", major: "y", minor: "y" }).map((x) => x.name)).toEqual(["A", "B", "C", "D"])
  })
})
