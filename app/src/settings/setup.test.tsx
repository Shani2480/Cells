import { act, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { beforeEach, describe, expect, it, vi } from "vitest"
import type { Platform, SettingsPort } from "@cells/platform"
import { App } from "@/App"
import { PlatformContext } from "@/platform/context"
import { SheetController } from "@/sheet/controller"
import { MemoryFiles } from "@/test/memory-files"
import { useSettings } from "./settings"
import { DEFAULT_SETTINGS } from "./shared"

function renderWith(initial: string | null) {
  const file = { text: initial }
  const settings: SettingsPort = {
    read: async () => file.text,
    write: async (text) => {
      file.text = text
    },
    location: async () => "~/.config/inkspire/settings.json",
    systemFacts: async () => ({ user: "alex", full_name: "Alex Demo", lang: "en_GB.UTF-8" }),
  }
  const platform: Platform = {
    os: "linux",
    clipboard: { write: vi.fn(async () => {}), read: vi.fn(async () => ({ text: "" })) },
    window: {
      controllable: false,
      minimize: async () => {},
      toggleMaximize: async () => {},
      close: async () => {},
      isMaximized: async () => false,
      setTitle: async () => {},
      onMaximizedChange: () => () => {},
      onCloseRequested: () => () => {},
    },
    files: new MemoryFiles(),
    settings,
    opens: { onRequest: () => () => {} },
  }
  const controller = new SheetController()
  render(
    <PlatformContext.Provider value={platform}>
      <App controller={controller} />
    </PlatformContext.Provider>,
  )
  return { file, controller }
}

const saved = (file: { text: string | null }) => JSON.parse(file.text ?? "{}")

beforeEach(() => useSettings.setState({ ...DEFAULT_SETTINGS, loaded: false }))

describe("first-run setup", () => {
  it("starts from the system, and Finish saves it for both apps", async () => {
    const { file, controller } = renderWith(null)
    expect(await screen.findByText("Welcome to Cells")).toBeInTheDocument()
    expect(screen.getByLabelText("User name:")).toHaveValue("Alex Demo")
    expect(screen.getByLabelText("Initials:")).toHaveValue("AD")
    fireEvent.click(screen.getByRole("button", { name: "Next" }))
    expect(screen.getByLabelText("Region:")).toHaveValue("en-GB")
    expect(screen.getByLabelText("Sample date")).toHaveTextContent("31/01/2026")
    fireEvent.click(screen.getByRole("button", { name: "Next" }))
    fireEvent.click(screen.getByRole("radio", { name: "Black" }))
    await waitFor(() => expect(document.documentElement).toHaveClass("dark"))
    fireEvent.click(screen.getByRole("button", { name: "Finish" }))
    await waitFor(() => expect(saved(file).apps?.cells?.setupDone).toBe(true))
    expect(saved(file)).toMatchObject({
      user: { name: "Alex Demo", initials: "AD" },
      regional: { locale: "en-GB", dateFormat: "dd/MM/yyyy" },
      appearance: { theme: "dark" },
    })
    expect(screen.queryByText("Welcome to Cells")).not.toBeInTheDocument()
    expect(controller.annotations.author).toBe("Alex Demo")
    act(() => void useSettings.setState({ theme: "system" }))
  })

  it("says what came from Inkspire, and Skip keeps it", async () => {
    const { file } = renderWith(
      JSON.stringify({ user: { name: "Sam Example", initials: "SE" }, apps: { inkspire: { setupDone: true } } }),
    )
    expect(await screen.findByText(/Filled in from the settings you share with Inkspire/)).toBeInTheDocument()
    expect(screen.getByLabelText("User name:")).toHaveValue("Sam Example")
    fireEvent.click(screen.getByRole("button", { name: "Skip" }))
    await waitFor(() => expect(saved(file).apps?.cells?.setupDone).toBe(true))
    expect(saved(file).user.name).toBe("Sam Example")
    expect(saved(file).apps.inkspire).toEqual({ setupDone: true })
  })

  it("doesn't show once it's done", async () => {
    renderWith(JSON.stringify({ apps: { cells: { setupDone: true } } }))
    await waitFor(() => expect(useSettings.getState().loaded).toBe(true))
    expect(screen.queryByText("Welcome to Cells")).not.toBeInTheDocument()
  })
})

describe("File › Options", () => {
  it("changes the user name and separators, saving only what changed", async () => {
    const { file } = renderWith(
      JSON.stringify({
        user: { name: "Sam Example", initials: "SE" },
        appearance: { theme: "light" },
        apps: { cells: { setupDone: true } },
      }),
    )
    await waitFor(() => expect(useSettings.getState().loaded).toBe(true))
    fireEvent.click(screen.getByRole("button", { name: "File" }))
    fireEvent.click(await screen.findByRole("button", { name: "Options" }))
    expect(screen.getByText("Excel Options")).toBeInTheDocument()
    expect(
      await screen.findByText(/shared with Inkspire \(~\/\.config\/inkspire\/settings\.json\)/),
    ).toBeInTheDocument()
    fireEvent.change(screen.getByLabelText("User name:"), { target: { value: "Sam Q. Example" } })
    // Inkspire switches to dark meanwhile.
    file.text = JSON.stringify({ ...saved(file), appearance: { theme: "dark" } })
    fireEvent.click(screen.getByRole("button", { name: "Advanced" }))
    fireEvent.change(screen.getByLabelText("Decimal separator:"), { target: { value: "," } })
    fireEvent.change(screen.getByLabelText("Thousands separator:"), { target: { value: "," } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(screen.getByRole("alert")).toHaveTextContent("must be different")
    fireEvent.change(screen.getByLabelText("Thousands separator:"), { target: { value: "." } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    await waitFor(() => expect(saved(file).user.name).toBe("Sam Q. Example"))
    // The initials followed the name, since they hadn't been typed over.
    expect(saved(file).user.initials).toBe("SQE")
    expect(saved(file).regional).toMatchObject({ decimalSeparator: ",", groupSeparator: "." })
    expect(saved(file).appearance.theme).toBe("dark")
  })
})
