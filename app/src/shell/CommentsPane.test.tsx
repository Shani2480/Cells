import { fireEvent, render, screen, within } from "@testing-library/react"
import { act } from "react"
import { describe, expect, it } from "vitest"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { SheetController } from "@/sheet/controller"
import { CommentsPane } from "./CommentsPane"

function setup() {
  const c = new SheetController()
  c.annotations.author = "Alex Demo"
  const comment = (r: number, col: number, text: string) => {
    c.goTo({ r0: r, c0: col, r1: r, c1: col })
    c.annotations.newComment()
    c.annotations.post(text)
    c.annotations.closeThread()
  }
  comment(4, 1, "Check the total")
  comment(1, 2, "Source?")
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  return c
}

describe("the Comments pane", () => {
  it("Show Comments toggles it and shows pressed", () => {
    const c = setup()
    const registry = buildRegistry()
    const env = { c } as unknown as RibbonEnv
    expect(registry.isBuilt("ReviewShowAllComments")).toBe(true)
    registry.get("ReviewShowAllComments")!.run!(env)
    expect(c.annotations.paneOpen).toBe(true)
    expect(registry.get("ReviewShowAllComments")!.pressed!(env)).toBe(true)
  })

  it("lists the threads in reading order; picking one goes to its cell and opens it", () => {
    const c = setup()
    c.annotations.togglePane(true)
    render(<CommentsPane controller={c} onDone={() => undefined} />)
    const pane = screen.getByRole("complementary", { name: "Comments" })
    const threads = within(pane).getAllByRole("button", { name: /^Comment on/ })
    expect(threads.map((t) => t.getAttribute("aria-label"))).toEqual(["Comment on C2", "Comment on B5"])
    expect(threads[1]).toHaveTextContent("Check the total")
    expect(threads[1]).toHaveTextContent("AD")
    fireEvent.click(threads[1])
    expect(c.selection.active).toEqual({ row: 4, col: 1 })
    expect(c.annotations.open).toMatchObject({ row: 4, col: 1 })
  })

  it("New starts a thread on the active cell; Close and Esc close it", () => {
    const c = setup()
    c.annotations.togglePane(true)
    render(<CommentsPane controller={c} onDone={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "New" }))
    expect(c.annotations.open).toEqual({ row: 0, col: 0, draft: true })
    fireEvent.keyDown(screen.getByRole("complementary", { name: "Comments" }), { key: "Escape" })
    expect(c.annotations.paneOpen).toBe(false)
    act(() => c.annotations.togglePane(true))
    fireEvent.click(screen.getByRole("button", { name: "Close" }))
    expect(c.annotations.paneOpen).toBe(false)
  })
})
