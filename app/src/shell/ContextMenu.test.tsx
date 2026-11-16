import { fireEvent, render, screen, within } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { ContextMenu, type MenuItem } from "./ContextMenu"

describe("the context menu's submenus", () => {
  const setup = () => {
    const picked = vi.fn()
    const onClose = vi.fn()
    const items: MenuItem[] = [
      { kind: "item", label: "Copy", onSelect: () => picked("Copy") },
      {
        kind: "item",
        label: "Sort",
        onSelect: () => undefined,
        items: [
          { kind: "item", label: "Sort A to Z", onSelect: () => picked("A to Z") },
          { kind: "item", label: "Sort Z to A", onSelect: () => picked("Z to A") },
        ],
      },
    ]
    render(<ContextMenu x={10} y={10} items={items} onClose={onClose} />)
    return { picked, onClose }
  }

  it("opens on hover, and a choice in it runs and closes the menu", () => {
    const { picked, onClose } = setup()
    const sort = screen.getByRole("menuitem", { name: "Sort" })
    expect(sort).toHaveAttribute("aria-haspopup", "menu")
    fireEvent.pointerEnter(sort)
    const sub = screen.getByRole("menu", { name: "Sort" })
    fireEvent.click(within(sub).getByRole("menuitem", { name: "Sort Z to A" }))
    expect(picked).toHaveBeenCalledWith("Z to A")
    expect(onClose).toHaveBeenCalled()
  })

  it("opens with a click or Right and goes back with Left; hovering another row closes it", () => {
    const { onClose } = setup()
    const sort = screen.getByRole("menuitem", { name: "Sort" })
    fireEvent.click(sort)
    expect(onClose).not.toHaveBeenCalled()
    expect(screen.getByRole("menu", { name: "Sort" })).toBeInTheDocument()
    fireEvent.pointerEnter(screen.getByRole("menuitem", { name: "Copy" }))
    expect(screen.queryByRole("menu", { name: "Sort" })).toBeNull()
    sort.focus()
    fireEvent.keyDown(sort, { key: "ArrowRight" })
    const sub = screen.getByRole("menu", { name: "Sort" })
    const first = within(sub).getByRole("menuitem", { name: "Sort A to Z" })
    first.focus()
    fireEvent.keyDown(first, { key: "ArrowLeft" })
    expect(screen.queryByRole("menu", { name: "Sort" })).toBeNull()
    expect(document.activeElement).toBe(sort)
  })
})
