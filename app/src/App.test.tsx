import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import type { Platform } from "@cells/platform"
import { App } from "./App"
import { PlatformContext } from "./platform/context"
import { SheetController } from "./sheet/controller"
import { MemoryFiles } from "./test/memory-files"
import { MemoryLinks } from "./test/memory-links"
import { MemoryAi } from "./test/memory-ai"

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
  settings: {
    read: async () => JSON.stringify({ version: 1, apps: { cells: { setupDone: true } } }),
    write: async () => {},
    location: async () => "",
    systemFacts: async () => ({}),
  },
  links: new MemoryLinks(),
  ai: new MemoryAi(),
  opens: {
    onRequest(handler) {
      requests.handler = handler
      return () => {
        requests.handler = null
      }
    },
    external: async (target) => void externals.push(target),
  },
}

/** What links asked the system to open (Insert › Link). */
const externals: string[] = []

/** What the system would hand the window (cells:// links, files). */
const requests: { handler: ((items: string[]) => void) | null } = { handler: null }

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

  it("a cells:// link opens the workbook at its range, and again only moves there", async () => {
    renderApp()
    const { Workbook } = await import("@cells/core")
    const { saveXlsx } = await import("@cells/io")
    const wb = new Workbook(["Sales", "Costs"])
    wb.sheets[1].cells.set(4, 2, { value: 7 })
    files.disk.set("/home/alex/Q1 Budget.xlsx", await saveXlsx(wb, { fileName: "Q1 Budget.xlsx" }))
    const link = "cells://open?file=%2Fhome%2Falex%2FQ1%20Budget.xlsx&range=Costs!R5C3:R6C4"
    await act(async () => requests.handler?.([link]))
    expect((await screen.findAllByText("Q1 Budget - Cells", {}, { timeout: 5000 })).length).toBeGreaterThan(0)
    await waitFor(() => expect(screen.getByDisplayValue("C5")).toBeInTheDocument())
    expect(screen.getByRole("tab", { name: "Costs", selected: true })).toBeInTheDocument()
    const reads = files.reads
    await act(async () => requests.handler?.(["cells://open?file=%2Fhome%2Falex%2FQ1%20Budget.xlsx&range=Sales!B2"]))
    await waitFor(() => expect(screen.getByDisplayValue("B2")).toBeInTheDocument())
    expect(files.reads).toBe(reads)
  })

  it("Edit Data opens a chart's workbook from Inkspire, named after the document", async () => {
    renderApp()
    const { Workbook } = await import("@cells/core")
    const { saveXlsx } = await import("@cells/io")
    const wb = new Workbook()
    wb.sheets[0].cells.set(0, 1, { value: 10 })
    const path = "/home/alex/.cache/inkspire/edit/Report-Chart 1.xlsx"
    files.disk.set(path, await saveXlsx(wb, { fileName: "Report-Chart 1.xlsx" }))
    const link = `cells://edit?file=${encodeURIComponent(path)}&for=Report.docx`
    await act(async () => requests.handler?.([link]))
    expect((await screen.findAllByText("Chart in Report.docx - Cells", {}, { timeout: 5000 })).length).toBeGreaterThan(
      0,
    )
  })

  it("Edit Data for a file outside Inkspire's edit folder only opens it", async () => {
    renderApp()
    const { Workbook } = await import("@cells/core")
    const { saveXlsx } = await import("@cells/io")
    files.disk.set("/home/alex/Sales.xlsx", await saveXlsx(new Workbook(), { fileName: "Sales.xlsx" }))
    await act(async () => requests.handler?.(["cells://edit?file=%2Fhome%2Falex%2FSales.xlsx&for=Report.docx"]))
    expect((await screen.findAllByText("Sales - Cells", {}, { timeout: 5000 })).length).toBeGreaterThan(0)
  })

  it("a CSV dropped on the window opens", async () => {
    renderApp()
    files.disk.set("/home/alex/Dropped.csv", new TextEncoder().encode("a,b\n1,2\n"))
    act(() => files.drop("/home/alex/Dropped.csv"))
    expect((await screen.findAllByText("Dropped - Cells", {}, { timeout: 5000 })).length).toBeGreaterThan(0)
    expect(screen.getByRole("tab", { name: "Dropped" })).toBeInTheDocument()
  })
})

describe("the shell by screen", () => {
  const realMatch = window.matchMedia
  const size = (width: number, height: number) => {
    Object.defineProperty(window, "innerWidth", { configurable: true, value: width })
    Object.defineProperty(window, "innerHeight", { configurable: true, value: height })
  }
  const touch = () => {
    window.matchMedia = ((query: string) => ({
      ...realMatch(query),
      matches: query === "(pointer: coarse)",
    })) as typeof window.matchMedia
  }
  const restore = () => {
    window.matchMedia = realMatch
    size(1024, 768)
    document.documentElement.classList.remove("is-phone")
  }

  it("a phone gets the phone shell: no ribbon, dialogs full screen", () => {
    touch()
    size(390, 844)
    try {
      renderApp()
      expect(screen.getByRole("navigation", { name: "Sheets and tools" })).toBeInTheDocument()
      expect(screen.queryByRole("tablist", { name: "Ribbon tabs" })).toBeNull()
      expect(document.documentElement).toHaveClass("is-phone")
    } finally {
      restore()
    }
  })

  it("a long press (or right-click) on a phone opens the action sheet, not the desktop menu", async () => {
    touch()
    size(390, 844)
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
    Element.prototype.getBoundingClientRect = () => new DOMRect(0, 0, 390, 600)
    try {
      const c = renderApp()
      act(() => {
        c.wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: { value: "x" } }] }], {
          label: "Set",
        })
        c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
      })
      const canvas = document.querySelector(".grid-canvas") as HTMLCanvasElement
      fireEvent.contextMenu(canvas, { clientX: 120, clientY: 60 })
      expect(screen.queryByRole("menu")).toBeNull()
      const sheet = screen.getByRole("dialog", { name: "Cells" })
      const quick = within(sheet).getByRole("group", { name: "Clipboard" })
      expect(
        within(quick)
          .getAllByRole("button")
          .map((b) => b.textContent),
      ).toEqual(["Cut", "Copy", "Paste"])
      // Built items only, without shortcuts.
      expect(within(sheet).queryByRole("button", { name: /Quick Analysis/ })).toBeNull()
      expect(within(sheet).getByRole("button", { name: "Format Cells…" })).toBeInTheDocument()
      // Back (Esc to what has focus) closes it: the sheet took focus from the grid, which keeps Esc for itself.
      expect(sheet.contains(document.activeElement)).toBe(true)
      fireEvent.keyDown(document.activeElement!, { key: "Escape" })
      await waitFor(() => expect(screen.queryByRole("dialog", { name: "Cells" })).toBeNull())
      fireEvent.contextMenu(canvas, { clientX: 120, clientY: 60 })
      const again = screen.getByRole("dialog", { name: "Cells" })
      fireEvent.click(within(again).getByRole("button", { name: "Clear Contents" }))
      await waitFor(() => expect(screen.queryByRole("dialog", { name: "Cells" })).toBeNull())
      expect(c.sheet.cells.input(0, 0)).toBeUndefined()
    } finally {
      globalThis.ResizeObserver = saved.observer
      Element.prototype.getBoundingClientRect = saved.rect
      restore()
    }
  })

  it("a tablet keeps the ribbon, and a split screen switches live without losing the workbook", () => {
    touch()
    size(1280, 800)
    try {
      const c = renderApp()
      expect(screen.getByRole("tablist", { name: "Ribbon tabs" })).toBeInTheDocument()
      act(() => c.sheets.insert(1))
      act(() => {
        size(500, 800)
        window.dispatchEvent(new Event("resize"))
      })
      expect(screen.queryByRole("tablist", { name: "Ribbon tabs" })).toBeNull()
      expect(screen.getByRole("button", { name: "Sheet: Sheet2" })).toBeInTheDocument()
      act(() => {
        size(1280, 800)
        window.dispatchEvent(new Event("resize"))
      })
      expect(screen.getByRole("tab", { name: "Sheet2" })).toBeInTheDocument()
      expect(document.documentElement).not.toHaveClass("is-phone")
    } finally {
      restore()
    }
  })
})
