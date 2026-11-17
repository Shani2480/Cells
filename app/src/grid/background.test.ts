import { afterEach, describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { drawBackground, onBackgroundLoaded, type BackgroundPainter } from "./background"
import { computeLayout } from "./layout"

/** An Image that "loads" on the next tick with a 40 × 30 picture. */
class FakeImage {
  naturalWidth = 40
  naturalHeight = 30
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  set src(_: string) {
    queueMicrotask(() => this.onload?.())
  }
}

function painter() {
  const calls: string[] = []
  const p: BackgroundPainter = {
    save: () => calls.push("save"),
    restore: () => calls.push("restore"),
    beginPath: () => {},
    rect: () => {},
    clip: () => {},
    translate: (x, y) => calls.push(`translate ${x},${y}`),
    fillRect: (x, y, w, h) => calls.push(`fill ${x},${y} ${w}x${h}`),
    createPattern: () => ({}) as CanvasPattern,
    fillStyle: "",
    globalAlpha: 1,
  }
  return { p, calls }
}

const measure = (text: string) => text.length * 7

afterEach(() => vi.unstubAllGlobals())

describe("sheet background", () => {
  it("Background and Delete Background undo, and the picture tiles from A1 once loaded", async () => {
    vi.stubGlobal("Image", FakeImage)
    const c = new SheetController()
    c.sheets.setBackground("data:image/png;base64,QQ==")
    expect(c.sheet.background).toBe("data:image/png;base64,QQ==")
    expect(c.wb.undoLabel).toBe("Sheet Background")

    const redraw = vi.fn()
    const off = onBackgroundLoaded(redraw)
    const layout = computeLayout(c, 800, 400, measure)
    const first = painter()
    drawBackground(first.p, c, layout.panes[0])
    expect(first.calls).toEqual([]) // still loading
    await Promise.resolve()
    expect(redraw).toHaveBeenCalled()
    off()

    // Scrolled down 3 rows: A1's corner sits 60 px above the cell area.
    c.scrollTo(0, 60)
    const scrolled = computeLayout(c, 800, 400, measure)
    const second = painter()
    drawBackground(second.p, c, scrolled.panes[0])
    const pane = scrolled.panes[0]
    expect(second.calls).toContain(`translate ${pane.cols[0].start},${pane.rows[0].start - 60}`)

    c.sheets.setBackground(null)
    expect(c.sheet.background).toBeUndefined()
    c.wb.undo()
    expect(c.sheet.background).toBe("data:image/png;base64,QQ==")
  })

  it("isn't drawn in Page Layout view", async () => {
    vi.stubGlobal("Image", FakeImage)
    const c = new SheetController()
    c.sheets.setBackground("data:image/png;base64,Qg==")
    drawBackground(painter().p, c, computeLayout(c, 800, 400, measure).panes[0])
    await Promise.resolve()
    c.setView("pageLayout")
    const { p, calls } = painter()
    drawBackground(p, c, computeLayout(c, 800, 400, measure).panes[0])
    expect(calls).toEqual([])
  })
})
