import { isEquation, OFFICE_THEME, plainText, type SheetShape } from "@cells/core"
import { describe, expect, it } from "vitest"
import { objectSvg } from "@/drawing/render"
import { SheetController } from "./controller"

const only = (c: SheetController) => c.sheet.objects[0] as SheetShape

describe("Insert › Equation", () => {
  it("puts a built-in equation in a text box that fits it, as one undo step", () => {
    const c = new SheetController()
    const shape = c.objects.insertEquation("x=(−b±√(b^2−4ac))/2a")
    expect(isEquation(shape)).toBe(true)
    expect(shape.name).toMatch(/^TextBox \d+$/)
    expect(c.wb.undoLabel).toBe("Insert Equation")
    const box = c.objects.box(shape)
    expect(box.w).toBeGreaterThan(60)
    expect(box.h).toBeGreaterThan(30)
    // Drawn built up: a fraction bar and a radical, no linear text.
    const svg = objectSvg(shape, box.w, box.h, OFFICE_THEME)
    expect(svg).toContain("<line")
    expect(svg).toMatch(/<path d="M[^"]+"[^>]*stroke-linejoin="round"/)
    expect(svg).not.toContain("^")
    c.wb.undo()
    expect(c.sheet.objects).toHaveLength(0)
  })

  it("a new equation opens for typing, shows Office's prompt while empty, and builds up what is typed", () => {
    const c = new SheetController()
    const shape = c.objects.insertEquation()
    expect(c.objects.editing).toBe(shape.id)
    const box = c.objects.box(shape)
    expect(objectSvg(only(c), box.w, box.h, OFFICE_THEME)).toContain("Type equation here.")
    c.objects.setText(shape.id, "a^2+b^2=c^2")
    expect(c.objects.editing).toBeNull()
    expect(only(c).text!.paragraphs[0]!.math![0]!.map((n) => n.t)).toEqual(["sSup", "r", "sSup", "r", "sSup"])
    // Shown again for typing in the linear form, letters in math italic as Office shows them.
    expect(plainText(only(c).text)).toBe("𝑎^2+𝑏^2=𝑐^2")
    expect(c.wb.undoLabel).toBe("Typing")
    // Two lines: two lines of one equation.
    c.objects.setText(shape.id, "x=1\ny=2")
    expect(only(c).text!.paragraphs).toHaveLength(1)
    expect(only(c).text!.paragraphs[0]!.math).toHaveLength(2)
  })

  it("the Equation tab types where the caret was, and typing goes on after it", () => {
    const c = new SheetController()
    const shape = c.objects.insertEquation("x+y")
    c.objects.select(shape.id)
    // The editor closed with the caret before "+y" (two characters from the end, 𝑦 being one surrogate pair).
    c.objects.lastCaret = { id: shape.id, fromEnd: 3 }
    c.objects.insertIntoEquation(shape.id, "^2")
    expect(plainText(only(c).text)).toBe("𝑥^2+𝑦")
    expect(c.objects.editing).toBe(shape.id)
    expect(c.objects.caret).toEqual({ id: shape.id, fromEnd: 3 })
    // With the editor open, the editor itself takes it at its caret.
    const typed: string[] = []
    c.objects.textInserter = (t) => typed.push(t)
    c.objects.insertIntoEquation(shape.id, "√□")
    expect(typed).toEqual(["√□"])
  })
})
