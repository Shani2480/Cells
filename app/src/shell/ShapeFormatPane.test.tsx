import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import type { SheetShape } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { ShapeFormatPane } from "./ShapeFormatPane"

/** Format Shape: each field changes the selected shapes as one undo step; Ctrl+1 opens it with a shape selected. */
function setup() {
  const c = new SheetController()
  const shape = c.objects.insertShape("rect", { x: 96, y: 96, w: 192, h: 96 })
  c.objects.setText(shape.id, "Plan")
  c.objects.select(shape.id)
  expect(c.keyReady({ key: "1", primary: true, shift: false, alt: false })).toBe(true)
  expect(c.objects.formatPaneOpen).toBe(true)
  const view = render(<ShapeFormatPane controller={c} onDone={() => undefined} />)
  const get = () => c.objects.find(shape.id) as SheetShape
  const type = (label: string, value: string) => {
    const input = screen.getByLabelText(label)
    fireEvent.change(input, { target: { value } })
    fireEvent.keyDown(input, { key: "Enter" })
    view.rerender(<ShapeFormatPane controller={c} onDone={() => undefined} />)
  }
  return { c, get, type, view }
}

describe("Format Shape pane", () => {
  it("Shape Options: fill transparency, line width and dash, size, rotation and placement", () => {
    const { c, get, type, view } = setup()
    type("Fill transparency", "25")
    expect(get().fill).toMatchObject({ kind: "solid", color: { alpha: 0.75 } })
    type("Line width", "3")
    expect(get().line.width).toBe(3)
    fireEvent.change(screen.getByLabelText("Dash type"), { target: { value: "dash" } })
    expect(get().line.dash).toBe("dash")
    type("Shape width", "3")
    expect(c.objects.box(get()).w).toBeCloseTo(288)
    type("Rotation", "-30")
    expect(get().rotation).toBe(330)
    fireEvent.click(screen.getByLabelText("Don't move or size with cells"))
    expect(get().placement).toBe("fixed")
    c.undo()
    expect(get().placement).toBeUndefined()
    fireEvent.click(screen.getByLabelText("No fill"))
    expect(get().fill).toEqual({ kind: "none" })
    view.unmount()
  })

  it("Text Options: where the text sits, its direction, margins, wrapping and the shape fitting it", () => {
    const { c, get, type, view } = setup()
    fireEvent.click(screen.getByRole("tab", { name: "Text Options" }))
    view.rerender(<ShapeFormatPane controller={c} onDone={() => undefined} />)
    fireEvent.change(screen.getByLabelText("Vertical alignment"), { target: { value: "b" } })
    fireEvent.change(screen.getByLabelText("Text direction"), { target: { value: "vert" } })
    type("Left margin", "0.5")
    fireEvent.click(screen.getByLabelText("Wrap text in shape"))
    fireEvent.click(screen.getByLabelText("Resize shape to fit text"))
    expect(get().text).toMatchObject({ anchor: "b", vertical: "vert", wrap: false, autoFit: true })
    expect(get().text!.insets!.l).toBeCloseTo(48)
    expect(get().text!.paragraphs[0].runs[0].text).toBe("Plan")
    view.unmount()
  })
})
