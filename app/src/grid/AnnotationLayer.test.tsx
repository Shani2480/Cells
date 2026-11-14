import { act, fireEvent, render, screen } from "@testing-library/react"
import { parseRange, selectRange } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { AnnotationLayer } from "./AnnotationLayer"
import { computeLayout } from "./layout"

const measure = (text: string) => text.length * 7

function show(c: SheetController, hover: { row: number; col: number } | null = null) {
  const layout = computeLayout(c, 800, 400, measure)
  return render(
    <AnnotationLayer
      controller={c}
      layout={layout}
      zoom={1}
      hover={hover}
      onKeepHover={() => undefined}
      onLeaveHover={() => undefined}
      onDone={() => undefined}
    />,
  )
}

function setup() {
  const c = new SheetController()
  c.annotations.author = "Alex Demo"
  c.setSelection(selectRange(parseRange("B2")!))
  return c
}

describe("note boxes", () => {
  it("a new note opens for typing with the author, and leaving it saves the text", () => {
    const c = setup()
    c.annotations.newNote()
    show(c)
    expect(screen.getByText("Alex Demo:")).toBeInTheDocument()
    const box = screen.getByLabelText("Note text")
    expect(box).toHaveFocus()
    fireEvent.change(box, { target: { value: "Budget check" } })
    act(() => box.blur())
    expect(c.sheet.annotations.get(1, 1)).toMatchObject({ text: "Budget check" })
  })

  it("shows a hidden note only while hovered", () => {
    const c = setup()
    c.annotations.newNote()
    c.annotations.finishNote("Hover me")
    const { rerender } = show(c)
    expect(screen.queryByText("Hover me")).toBeNull()
    const layout = computeLayout(c, 800, 400, measure)
    rerender(
      <AnnotationLayer
        controller={c}
        layout={layout}
        zoom={1}
        hover={{ row: 1, col: 1 }}
        onKeepHover={() => undefined}
        onLeaveHover={() => undefined}
        onDone={() => undefined}
      />,
    )
    expect(screen.getByText("Hover me")).toBeInTheDocument()
  })
})

describe("comment card", () => {
  it("starts a conversation, posts with Ctrl+Enter and then offers a reply", () => {
    const c = setup()
    c.annotations.newComment()
    const { rerender } = show(c)
    const input = screen.getByLabelText("Start a conversation")
    expect(input).toHaveFocus()
    fireEvent.change(input, { target: { value: "Why so high?" } })
    fireEvent.keyDown(input, { key: "Enter", ctrlKey: true })
    expect(c.sheet.annotations.get(1, 1)).toMatchObject({ kind: "comment", posts: [{ text: "Why so high?" }] })
    rerender(
      <AnnotationLayer
        controller={c}
        layout={computeLayout(c, 800, 400, measure)}
        zoom={1}
        hover={null}
        onKeepHover={() => undefined}
        onLeaveHover={() => undefined}
        onDone={() => undefined}
      />,
    )
    expect(screen.getByText("Why so high?")).toBeInTheDocument()
    expect(screen.getByLabelText("Reply")).toBeInTheDocument()
    expect(screen.getByText("AD")).toBeInTheDocument()
  })

  it("the thread menu deletes the thread", () => {
    const c = setup()
    c.annotations.newComment()
    c.annotations.post("Hello")
    show(c)
    fireEvent.click(screen.getByRole("button", { name: "More thread actions" }))
    fireEvent.click(screen.getByRole("menuitem", { name: "Delete thread" }))
    expect(c.sheet.annotations.size).toBe(0)
  })
})
