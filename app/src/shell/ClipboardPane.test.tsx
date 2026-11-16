import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { SheetController } from "@/sheet/controller"
import { ClipboardPane, clipPreview } from "./ClipboardPane"

function setup() {
  const c = new SheetController()
  const put = (row: number, col: number, text: string) => {
    c.goTo({ r0: row, c0: col, r1: row, c1: col })
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  put(0, 0, "Rent")
  put(0, 1, "1200")
  put(1, 0, "Food")
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 1 })
  c.copy(false)
  c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
  c.copy(false)
  c.copy(false)
  return c
}

describe("the Office Clipboard", () => {
  it("keeps what was copied, newest first, once each, up to 24", () => {
    const c = setup()
    expect(c.clipboardHistory.map((i) => clipPreview(i.text))).toEqual(["Food", "Rent  1200"])
    expect(clipPreview("a\tb\n\t\nc\n")).toBe("a  b\nc")
    for (let i = 0; i < 30; i++) {
      c.goTo({ r0: 5 + i, c0: 0, r1: 5 + i, c1: 0 })
      c.beginEdit("enter", `n${i}`)
      c.commitEdit()
      c.copy(false)
    }
    expect(c.clipboardHistory).toHaveLength(24)
    expect(clipPreview(c.clipboardHistory[0].text)).toBe("n29")
  })

  it("the Clipboard launcher opens the pane; an item pastes, Paste All pastes all, Clear All empties", () => {
    const c = setup()
    const registry = buildRegistry()
    registry.get("ShowClipboard")!.run!({ c } as unknown as RibbonEnv)
    expect(c.clipboardPaneOpen).toBe(true)
    const paste = vi.fn()
    render(<ClipboardPane controller={c} onPaste={paste} onDone={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "Paste Rent 1200" }))
    expect(paste).toHaveBeenLastCalledWith(c.clipboardHistory[1].text, c.clipboardHistory[1].html)
    fireEvent.click(screen.getByRole("button", { name: "Paste All" }))
    expect(paste).toHaveBeenLastCalledWith("Rent\t1200\nFood", "")
    fireEvent.click(screen.getAllByRole("button", { name: "Delete" })[0])
    expect(c.clipboardHistory).toHaveLength(1)
    fireEvent.click(screen.getByRole("button", { name: "Clear All" }))
    expect(c.clipboardHistory).toHaveLength(0)
    expect(screen.getByText(/Copy something/)).toBeInTheDocument()
  })
})
