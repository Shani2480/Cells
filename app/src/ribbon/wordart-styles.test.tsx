import { isChart, type SheetShape } from "@cells/core"
import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { WORDART_STYLES } from "@/drawing/wordart"
import { buildRegistry } from "./commands"
import type { RibbonEnv } from "./registry"

describe("Shape Format › WordArt Styles", () => {
  it("gives the selected shapes a WordArt look, and Clear WordArt takes it off (one undo step each)", () => {
    const c = new SheetController()
    const o = c.objects.insertShape("rect")
    c.objects.setText(o.id, "Sale")
    const line = c.objects.insertShape("line")
    c.objects.select([o.id, line.id])
    const env = { c } as RibbonEnv
    const command = buildRegistry().get("WordArtStylesGallery")!
    expect(command.enabled?.(env)).toBe(true)
    const entry = command.menu![0]
    const open = () => {
      if (typeof entry === "string" || !("panel" in entry)) throw new Error("not a panel")
      return render(<>{entry.panel(env, () => undefined)}</>)
    }
    const shape = () => c.objects.find(o.id) as SheetShape
    let view = open()
    fireEvent.click(screen.getByRole("menuitem", { name: WORDART_STYLES[3].name }))
    expect(shape().wordArt).toEqual({ fill: WORDART_STYLES[3].fill, effects: WORDART_STYLES[3].effects })
    const l = c.objects.find(line.id)!
    expect(!isChart(l) && l.kind === "shape" && l.wordArt).toBeFalsy()
    view.unmount()
    view = open()
    fireEvent.click(screen.getByRole("menuitem", { name: "Clear WordArt" }))
    expect(shape().wordArt).toBeUndefined()
    c.undo()
    expect(shape().wordArt?.effects?.glow).toBeTruthy()
    view.unmount()
  })
})
