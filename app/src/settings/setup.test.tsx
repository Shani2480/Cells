import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react"
import { beforeEach, describe, expect, it, vi } from "vitest"
import type { Platform, SettingsPort } from "@cells/platform"
import { App } from "@/App"
import { PlatformContext } from "@/platform/context"
import { SheetController } from "@/sheet/controller"
import { MemoryFiles } from "@/test/memory-files"
import { MemoryLinks } from "@/test/memory-links"
import { MemoryAi } from "@/test/memory-ai"
import { useSettings } from "./settings"
import { DEFAULT_SETTINGS } from "./shared"

// Each test renders the whole window, which takes a few seconds while the full suite runs in parallel.
vi.setConfig({ testTimeout: 20_000 })

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
    links: new MemoryLinks(),
    ai: new MemoryAi(),
  }
  const controller = new SheetController()
  const ai = platform.ai as MemoryAi
  render(
    <PlatformContext.Provider value={platform}>
      <App controller={controller} />
    </PlatformContext.Provider>,
  )
  return { file, controller, ai }
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
    fireEvent.click(screen.getByRole("button", { name: "Next" }))
    // AI is opt-in: off, with nothing more to choose until it's turned on.
    expect(screen.getByRole("checkbox", { name: /Turn on the AI tools/ })).not.toBeChecked()
    expect(screen.queryByLabelText("Use:")).not.toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Finish" }))
    await waitFor(() => expect(saved(file).apps?.cells?.setupDone).toBe(true))
    expect(saved(file).apps.cells.ai).toBe(false)
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

describe("AI tools (Home › AI Tools)", () => {
  const on = (provider = "ollama") =>
    JSON.stringify({ ai: { provider }, apps: { cells: { setupDone: true, ai: true } } })

  const open = async (item: string) => {
    fireEvent.click(screen.getByRole("button", { name: "AI Tools" }))
    fireEvent.click(await screen.findByRole("menuitem", { name: item }))
  }

  it("while off, says so and takes you to File › Options › AI to turn them on", async () => {
    const { file, ai } = renderWith(JSON.stringify({ apps: { cells: { setupDone: true } } }))
    await waitFor(() => expect(useSettings.getState().loaded).toBe(true))
    await open("Formula from Words...")
    expect(screen.getByText(/The AI tools are off/)).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Options…" }))
    expect(await screen.findByText("AI tools")).toBeInTheDocument()
    fireEvent.click(screen.getByRole("checkbox", { name: /Turn on the AI tools/ }))
    fireEvent.change(screen.getByLabelText("Use:"), { target: { value: "ollama" } })
    expect(screen.getByLabelText("Address:")).toHaveAttribute("placeholder", "http://127.0.0.1:11434")
    expect(screen.getByText(/Runs on this computer/)).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Test Connection" }))
    expect(await screen.findByText("Connected.")).toBeInTheDocument()
    expect(ai.tested[0]).toMatchObject({ provider: "ollama", baseUrl: "http://127.0.0.1:11434" })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    await waitFor(() => expect(saved(file).apps.cells.ai).toBe(true))
    expect(saved(file).ai.provider).toBe("ollama")
  })

  it("writes a formula from words, and inserts it only when asked", async () => {
    const { controller, ai } = renderWith(on())
    await waitFor(() => expect(useSettings.getState().aiEnabled).toBe(true))
    const s = controller.wb.sheets[0]
    s.cells.set(0, 1, { value: 1200 })
    s.cells.set(1, 1, { value: 45 })
    act(() => controller.goTo({ r0: 2, c0: 1, r1: 2, c1: 1 }))
    ai.reply = '{"formula": "=SUM(B1:B2)", "explanation": "Adds the two costs."}'
    await open("Formula from Words...")
    expect(screen.getByText(/Sends the cells near B3 to Ollama \(on this computer\)\./)).toBeInTheDocument()
    fireEvent.change(screen.getByLabelText(/Describe what B3 should work out/), { target: { value: "total cost" } })
    fireEvent.click(screen.getByRole("button", { name: "Write Formula" }))
    expect(await screen.findByLabelText("Formula:")).toHaveValue("=SUM(B1:B2)")
    expect(screen.getByText("Adds the two costs.")).toBeInTheDocument()
    expect(ai.asked[0].prompt).toBe('In cell B3 of sheet "Sheet1": total cost')
    expect(ai.asked[0].target).toEqual({ provider: "ollama", model: undefined, baseUrl: "http://127.0.0.1:11434" })
    expect(s.cells.get(2, 1)).toBeUndefined()
    fireEvent.click(within(screen.getByRole("dialog")).getByRole("button", { name: "Insert" }))
    await waitFor(() => expect(s.cells.get(2, 1)?.input).toBe("=SUM(B1:B2)"))
  })

  it("says plainly when the reply isn't a formula Cells can read", async () => {
    const { ai } = renderWith(on())
    await waitFor(() => expect(useSettings.getState().aiEnabled).toBe(true))
    ai.reply = '{"formula": "=SUM(B1:", "explanation": ""}'
    await open("Formula from Words...")
    fireEvent.change(screen.getByLabelText(/should work out/), { target: { value: "total" } })
    fireEvent.click(screen.getByRole("button", { name: "Write Formula" }))
    expect(await screen.findByRole("alert")).toHaveTextContent("isn't a formula Cells can read")
    expect(within(screen.getByRole("dialog")).queryByRole("button", { name: "Insert" })).not.toBeInTheDocument()
  })

  it("suggests a chart and inserts it with its title; cloud providers say the cells leave the computer", async () => {
    const { controller, ai } = renderWith(on("anthropic"))
    await waitFor(() => expect(useSettings.getState().aiEnabled).toBe(true))
    const s = controller.wb.sheets[0]
    s.cells.set(0, 0, { value: "Item" })
    s.cells.set(0, 1, { value: "Cost" })
    s.cells.set(1, 0, { value: "Rent" })
    s.cells.set(1, 1, { value: 1200 })
    act(() => controller.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 }))
    ai.reply = '{"chart": "pie", "title": "Where the money goes", "why": "Shares of a whole."}'
    await open("Suggest a Chart")
    expect(await screen.findByLabelText("Suggestion")).toHaveTextContent("Pie")
    expect(screen.getByText(/to Anthropic, outside this computer/)).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "Insert Chart" }))
    await waitFor(() => expect(s.charts).toHaveLength(1))
    expect(s.charts[0].spec.title?.text).toBe("Where the money goes")
  })

  it("explains a formula as it streams in, or asks for a cell with one", async () => {
    const { controller, ai } = renderWith(on())
    await waitFor(() => expect(useSettings.getState().aiEnabled).toBe(true))
    await open("Explain Formula")
    expect(screen.getByText(/Select a cell with a formula/)).toBeInTheDocument()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    controller.wb.sheets[0].cells.set(0, 0, { value: 3, input: "=1+2" })
    ai.reply = "Adds 1 and 2."
    await open("Explain Formula")
    expect(await screen.findByLabelText("Answer")).toHaveTextContent("Adds 1 and 2.")
  })
})
