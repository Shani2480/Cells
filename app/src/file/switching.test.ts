import { describe, expect, it, vi } from "vitest"
import type { WindowPort } from "@cells/platform"
import { MemoryFiles } from "@/test/memory-files"
import { FileCommands, type FileUi } from "./commands"
import { newDocument } from "./document"
import { cycleEntries, switchEntries } from "./switching"

const ui = { show: () => {} } as unknown as FileUi

describe("Switch Windows across the desktop's windows", () => {
  it("lists every Cells window in order, this one's own windows in its place, and goes to the next with Ctrl+F6", async () => {
    const files = new FileCommands(new MemoryFiles(), ui, newDocument())
    files.doc.windows.newWindow()
    const focusWindow = vi.fn(async () => {})
    const win = {
      listWindows: async () => [
        { id: "main", title: "Budget.xlsx - Cells", current: false },
        { id: "book-2", title: "Book1 - Cells", current: true },
        { id: "book-3", title: "Trips.xlsx - Cells", current: false },
      ],
      focusWindow,
    } as unknown as WindowPort
    const list = await switchEntries(files, win)
    expect(list.map((e) => [e.title, e.current])).toEqual([
      ["Budget.xlsx", false],
      ["Book1:1", false],
      ["Book1:2", true],
      ["Trips.xlsx", false],
    ])
    await cycleEntries(files, win, 1)
    expect(focusWindow).toHaveBeenCalledWith("book-3")
    await cycleEntries(files, win, -1)
    expect(files.doc.windows.number(files.doc.windows.active)).toBe(1)
  })

  it("without other windows (Android, a browser) lists only this window's", async () => {
    const files = new FileCommands(new MemoryFiles(), ui, newDocument())
    expect((await switchEntries(files, {} as WindowPort)).map((e) => e.title)).toEqual(["Book1"])
  })
})
