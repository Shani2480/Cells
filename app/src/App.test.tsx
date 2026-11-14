import { act, fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import type { Platform } from "@cells/platform"
import { App } from "./App"
import { PlatformContext } from "./platform/context"
import { SheetController } from "./sheet/controller"

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
  },
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
})
