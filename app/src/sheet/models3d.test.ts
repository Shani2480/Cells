import { model3dXml, type Model3D, type SheetPicture } from "@cells/core"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"

// No WebGL here: the picture stands in as a model whose outline is 2:1 turned flat and 1:1 turned a quarter, at the
// view square's scale; `fit` scales it into the box, as the real renderer does.
vi.mock("@/drawing/model3d", () => ({
  renderModel: async (m: Model3D, o: { box?: { w: number; h: number }; fit?: boolean } = {}) => {
    const side = m.rot.y % 180 === 90 ? 0.5 : 1
    let vp = m.viewport
    let size = { w: vp * side * 0.8, h: vp * 0.4 }
    if (o.fit && o.box) {
      const k = Math.min(o.box.w / size.w, o.box.h / size.h)
      vp *= k
      size = { w: size.w * k, h: size.h * k }
    }
    return { src: `data:image/png;base64,${btoa(JSON.stringify(m.rot))}`, size, viewport: vp }
  },
}))

const model = (): Model3D => ({
  data: new Uint8Array([103, 108, 84, 70]),
  rot: { x: 0, y: 0, z: 0 },
  viewport: 288,
  xml: model3dXml({ meterPerUnit: 0.5, preTrans: [-2.5, 0, 0], cameraZ: 1.5 }),
})
const only = (c: SheetController) => c.sheet.objects[0] as SheetPicture

describe("Insert › 3D Models", () => {
  it("puts the model in at the active cell, its frame fitted to it, as one undo step", async () => {
    const c = new SheetController()
    const pic = await c.objects.insertModel3D(model())
    expect(pic?.name).toMatch(/^3D Model \d+$/)
    expect(c.wb.undoLabel).toBe("Insert 3D Model")
    const box = c.objects.box(only(c))
    expect(box.w).toBeCloseTo(230.4)
    expect(box.h).toBeCloseTo(115.2)
    expect(c.objects.model()?.id).toBe(pic!.id)
    c.wb.undo()
    expect(c.sheet.objects).toHaveLength(0)
  })

  it("turning draws it anew and fits the frame around the same middle; undo turns it back", async () => {
    const c = new SheetController()
    const pic = (await c.objects.insertModel3D(model()))!
    const before = c.objects.box(only(c))
    expect(await c.objects.setModelView(pic.id, { rot: { x: 0, y: -270, z: 0 } }, "3D Rotation")).toBe(true)
    // Kept between 0 and 360, as Office stores it.
    expect(only(c).model3d!.rot).toEqual({ x: 0, y: 90, z: 0 })
    expect(only(c).src).toContain(btoa(JSON.stringify({ x: 0, y: 90, z: 0 })))
    const after = c.objects.box(only(c))
    expect(after.w).toBeCloseTo(115.2)
    expect(after.x + after.w / 2).toBeCloseTo(before.x + before.w / 2)
    expect(after.y + after.h / 2).toBeCloseTo(before.y + before.h / 2)
    expect(c.wb.undoLabel).toBe("3D Rotation")
    c.wb.undo()
    expect(only(c).model3d!.rot).toEqual({ x: 0, y: 0, z: 0 })
    expect(c.objects.box(only(c)).w).toBeCloseTo(before.w)
  })

  it("Reset 3D Model keeps the frame and scales the model into it", async () => {
    const c = new SheetController()
    const pic = (await c.objects.insertModel3D(model()))!
    await c.objects.setModelView(pic.id, { rot: { x: 0, y: 90, z: 0 } }, "3D Rotation")
    const frame = c.objects.box(only(c))
    await c.objects.setModelView(pic.id, { rot: { x: 0, y: 0, z: 0 } }, "Reset 3D Model", true)
    expect(c.objects.box(only(c)).w).toBeCloseTo(frame.w)
    expect(c.objects.box(only(c)).h).toBeCloseTo(frame.h)
    // The flat model is twice as wide: it fits by its width, at half the view square.
    expect(only(c).model3d!.viewport).toBeCloseTo(144)
  })

  it("sizing the frame scales the model with it (by the side that fits)", async () => {
    const c = new SheetController()
    const pic = (await c.objects.insertModel3D(model()))!
    const box = c.objects.box(only(c))
    c.objects.place(new Map([[pic.id, { ...box, w: box.w * 2, h: box.h * 2 }]]), "Size")
    expect(only(c).model3d!.viewport).toBeCloseTo(576)
    expect(c.wb.undoLabel).toBe("Size")
    // The picture is drawn again for the new frame, not as a step of its own.
    await new Promise((r) => setTimeout(r, 0))
    expect(c.wb.undoLabel).toBe("Size")
  })
})
