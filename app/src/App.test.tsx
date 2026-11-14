import { act, fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import type { Platform } from "@cells/platform"
import { App } from "./App"
import { PlatformContext } from "./platform/context"
import { SheetController } from "./sheet/controller"
import { MemoryFiles } from "./test/memory-files"

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
}

function renderApp() {
  const controller = new SheetController()
  render(
    <PlatformContext.Provider value={platform}>
      <App controller={controller} />
    </PlatformContext.Provider>,
  )
  return controller
}

describe("App shell", () => {
  it("shows the title, Name Box and Ready status", () => {
    renderApp()
    expect(screen.getByText("Book1 - Cells")).toBeInTheDocument()
    expect(screen.getByDisplayValue("A1")).toBeInTheDocument()
    expect(screen.getByText("Ready")).toBeInTheDocument()
  })

  it("shows Sum, Average and Count for a numeric selection", () => {
    const c = renderApp()
    act(() => {
      c.wb.transact(
        [
          {
            kind: "setCells",
            sheet: 0,
            cells: [
              { row: 0, col: 0, data: { value: 2 } },
              { row: 1, col: 0, data: { value: 4 } },
            ],
          },
        ],
        { label: "Set" },
      )
      c.goTo({ r0: 0, c0: 0, r1: 1, c1: 0 })
    })
    expect(screen.getByText(/Average: 3/)).toBeInTheDocument()
    expect(screen.getByText(/Count: 2/)).toBeInTheDocument()
    expect(screen.getByText(/Sum: 6/)).toBeInTheDocument()
  })

  it("shows the average at a cell's width, as Excel does", () => {
    const c = renderApp()
    act(() => {
      c.wb.transact(
        [{ kind: "setCells", sheet: 0, cells: [0, 0, 1].map((v, row) => ({ row, col: 0, data: { value: v + 44 } })) }],
        { label: "Set" },
      )
      c.goTo({ r0: 0, c0: 0, r1: 2, c1: 0 })
    })
    expect(screen.getByText("Average: 44.33333333")).toBeInTheDocument()
  })

  it("the Name Box goes to a typed reference", () => {
    const c = renderApp()
    const box = screen.getByDisplayValue("A1")
    fireEvent.focus(box)
    fireEvent.change(box, { target: { value: "C7" } })
    fireEvent.keyDown(box, { key: "Enter" })
    expect(c.selection.active).toEqual({ row: 6, col: 2 })
  })

  it("Ctrl+G opens Go To, and Enter in it selects the reference", () => {
    const c = renderApp()
    act(() => {
      c.keyReady({ key: "g", primary: true, shift: false, alt: false })
    })
    const field = screen.getByLabelText(/Reference/)
    fireEvent.change(field, { target: { value: "B2:D4" } })
    fireEvent.keyDown(field, { key: "Enter" })
    expect(c.selection.ranges[0]).toEqual({ r0: 1, c0: 1, r1: 3, c1: 3 })
    expect(screen.queryByRole("dialog")).toBeNull()
  })

  it("the cell menu follows Excel's list, with unbuilt items disabled", () => {
    // Give the grid a size, so it lays out and a right-click hits a cell.
    const saved = { observer: globalThis.ResizeObserver, rect: Element.prototype.getBoundingClientRect }
    globalThis.ResizeObserver = class {
      constructor(private readonly callback: () => void) {}
      observe() {
        this.callback()
      }
      unobserve() {}
      disconnect() {}
    } as unknown as typeof ResizeObserver
    Element.prototype.getBoundingClientRect = () => new DOMRect(0, 0, 800, 500)
    try {
      renderApp()
    } finally {
      globalThis.ResizeObserver = saved.observer
      Element.prototype.getBoundingClientRect = saved.rect
    }
    const canvas = document.querySelector(".grid-canvas") as HTMLCanvasElement
    fireEvent.contextMenu(canvas, { clientX: 120, clientY: 60 })
    const menu = screen.queryByRole("menu")
    expect(menu).not.toBeNull()
    if (!menu) return
    const labels = [...menu.querySelectorAll('[role="menuitem"]')].map((b) => b.textContent)
    expect(labels).toContain("Format Cells...")
    expect(labels).toContain("New Note")
    expect(screen.getByRole("menuitem", { name: "New Note" })).toBeEnabled()
    expect(screen.getByRole("menuitem", { name: "Define Name..." })).toBeEnabled()
    expect(screen.getByRole("menuitem", { name: "Quick Analysis" })).toBeDisabled()
    expect(screen.getByRole("menuitem", { name: "Format Cells..." })).toBeEnabled()
  })
})

describe("File commands in the window", () => {
  const files = platform.files as MemoryFiles

  function edit(c: SheetController, text: string) {
    act(() => {
      c.beginEdit("enter", text)
      c.commitEdit()
    })
  }

  it("Ctrl+S on a new workbook asks where, then the title shows the file's name", async () => {
    const c = renderApp()
    edit(c, "42")
    files.nextSave = "/home/alex/Answers.xlsx"
    await act(async () => {
      fireEvent.keyDown(window, { key: "s", ctrlKey: true })
    })
    await vi.waitFor(() => expect(files.disk.has("/home/alex/Answers.xlsx")).toBe(true))
    await vi.waitFor(() => expect(screen.getByText("Answers - Cells")).toBeInTheDocument())
  })

  it("Ctrl+N with unsaved changes asks to save; Don't Save opens a new workbook", async () => {
    const c = renderApp()
    edit(c, "7")
    await act(async () => {
      fireEvent.keyDown(window, { key: "n", ctrlKey: true })
    })
    expect(await screen.findByText(/Want to save your changes to/)).toBeInTheDocument()
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Don't Save" }))
    })
    await vi.waitFor(() => expect(screen.queryByText(/Want to save/)).toBeNull())
    expect(screen.getAllByText("Book2 - Cells").length).toBeGreaterThan(0)
    expect(screen.getByDisplayValue("A1")).toBeInTheDocument()
  })

  it("Ctrl+O opens a workbook in place of the untouched one", async () => {
    renderApp()
    const { Workbook } = await import("@cells/core")
    const { saveXlsx } = await import("@cells/io")
    const wb = new Workbook()
    wb.sheets[0].cells.set(0, 0, { value: "Hello" })
    files.disk.set("/home/alex/Hello.xlsx", await saveXlsx(wb, { fileName: "Hello.xlsx" }))
    files.nextOpen = "/home/alex/Hello.xlsx"
    await act(async () => {
      fireEvent.keyDown(window, { key: "o", ctrlKey: true })
    })
    expect((await screen.findAllByText("Hello - Cells", {}, { timeout: 5000 })).length).toBeGreaterThan(0)
    expect(screen.getByDisplayValue("A1")).toBeInTheDocument()
  })

  it("File opens on Home, with the Recent list", async () => {
    renderApp()
    fireEvent.click(screen.getByRole("button", { name: "File" }))
    expect(screen.getByRole("dialog", { name: "File" })).toBeInTheDocument()
    expect(screen.getByRole("heading", { level: 1 }).textContent).toMatch(/^Good (morning|afternoon|evening)$/)
    expect(screen.getByRole("button", { name: "Blank workbook" })).toBeInTheDocument()
    expect(screen.getByRole("list", { name: "Recent files" })).toBeInTheDocument()
  })
})
