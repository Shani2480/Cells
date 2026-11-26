import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { drawCells, type CellPainter } from "./draw-cells"
import { computeLayout } from "./layout"
import { LIGHT_THEME } from "./theme"

/** Records the colour of every filled rectangle, text and line drawn. */
function recorder() {
  const used: { what: string; color: string }[] = []
  const noop = () => {}
  const painter = {
    save: noop,
    restore: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    moveTo: noop,
    lineTo: noop,
    setLineDash: noop,
    fill: noop,
    arc: noop,
    closePath: noop,
    translate: noop,
    rotate: noop,
    setTransform: noop,
    createLinearGradient: () => ({ addColorStop: noop }) as unknown as CanvasGradient,
    fillStyle: "" as string | CanvasGradient,
    strokeStyle: "" as string | CanvasGradient,
    lineWidth: 1,
    font: "",
    textAlign: "left" as CanvasTextAlign,
    textBaseline: "alphabetic" as CanvasTextBaseline,
    lineDashOffset: 0,
    globalAlpha: 1,
    fillRect() {
      used.push({ what: "fill", color: String(this.fillStyle) })
    },
    strokeRect() {
      used.push({ what: "line", color: String(this.strokeStyle) })
    },
    stroke() {
      used.push({ what: "line", color: String(this.strokeStyle) })
    },
    fillText() {
      used.push({ what: "text", color: String(this.fillStyle) })
    },
  }
  return { painter: painter as unknown as CellPainter, used }
}

describe("printing in black and white (Page Setup › Sheet)", () => {
  it("leaves fills out and draws text and borders in black, as Excel prints them", () => {
    const c = new SheetController()
    const styles = c.wb.styles
    const red = styles.intern({
      ...styles.get(0),
      fill: "#FF0000",
      fontColor: "#0000FF",
      borderBottom: { style: "thin", color: "#00FF00" },
    })
    c.sheet.cells.set(0, 0, { value: "Filled", style: red })
    c.showGridlines = false
    const layout = computeLayout(c, 400, 200, (t) => t.length * 7)
    const run = (mono: boolean) => {
      const { painter, used } = recorder()
      drawCells({ painter, layout, controller: c, theme: LIGHT_THEME, measure: (t) => t.length * 7, mono })
      return used
    }
    const colour = run(false)
    expect(colour).toContainEqual({ what: "fill", color: "#FF0000" })
    expect(colour).toContainEqual({ what: "text", color: "#0000FF" })
    const mono = run(true)
    expect(mono.filter((u) => u.what === "fill")).toEqual([])
    expect(mono.filter((u) => u.what === "text").map((u) => u.color)).toEqual(["#000000"])
    expect(mono.some((u) => u.what === "line" && u.color === "#000000")).toBe(true)
    expect(mono.some((u) => u.what === "line" && u.color === "#00FF00")).toBe(false)
  })
})
