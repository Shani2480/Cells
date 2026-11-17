import { isChart, OFFICE_THEME, type SheetPicture } from "@cells/core"
import { fireEvent, render, screen, waitFor } from "@testing-library/react"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { objectSvg } from "@/drawing/render"
import { SheetController } from "@/sheet/controller"
import { SignatureSetupDialog } from "./SignatureSetupDialog"

const line = (c: SheetController) =>
  c.sheet.objects.find((o): o is SheetPicture => !isChart(o) && o.kind === "picture" && !!o.signature)

/** An Image that loads at once (jsdom has no canvas: the SVG then stands in for the PNG). */
class FakeImage {
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  set src(_: string) {
    queueMicrotask(() => this.onload?.())
  }
}

beforeEach(() => vi.stubGlobal("Image", FakeImage))
afterEach(() => vi.unstubAllGlobals())

describe("Insert › Signature Line", () => {
  it("Signature Setup puts a line at the active cell, drawn with the signer; setting it up again is one undo step", async () => {
    const c = new SheetController()
    let closed = 0
    const { unmount } = render(<SignatureSetupDialog controller={c} onClose={() => closed++} />)
    fireEvent.change(screen.getByLabelText(/Suggested signer \(for example/), { target: { value: "Sam Sample" } })
    fireEvent.change(screen.getByLabelText(/title/), { target: { value: "Director" } })
    expect((screen.getByLabelText(/Instructions to the signer/) as HTMLTextAreaElement).value).toMatch(
      /^Before signing/,
    )
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    await waitFor(() => expect(closed).toBe(1))
    unmount()

    const added = line(c)!
    expect(added.signature).toMatchObject({ signer: "Sam Sample", title: "Director", showDate: true })
    expect(added.signature!.id).toMatch(/^\{[0-9A-F-]{36}\}$/)
    expect(c.wb.undoLabel).toBe("Insert Signature Line")
    const svg = objectSvg(added, 192, 96, OFFICE_THEME)
    expect(svg).toContain(">X</text>")
    expect(svg).toContain(">Sam Sample</text>")
    expect(svg).toContain(">Director</text>")

    render(<SignatureSetupDialog controller={c} editing={added.id} onClose={() => closed++} />)
    expect((screen.getByLabelText(/Suggested signer \(for example/) as HTMLInputElement).value).toBe("Sam Sample")
    fireEvent.change(screen.getByLabelText(/Suggested signer \(for example/), { target: { value: "Alex Example" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    await waitFor(() => expect(closed).toBe(2))
    expect(line(c)!.signature).toMatchObject({ signer: "Alex Example", id: added.signature!.id })
    expect(c.wb.undoLabel).toBe("Signature Setup")
    c.wb.undo()
    expect(line(c)!.signature!.signer).toBe("Sam Sample")
  })
})
