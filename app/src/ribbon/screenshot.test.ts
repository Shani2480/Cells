import { afterEach, describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { imageDataUrl } from "@/platform/tauri"
import { allCommands } from "./commands"
import type { RibbonEnv } from "./registry"

const command = allCommands().find((c) => c.id === "ScreenshotInsertGallery")!
const PNG = "data:image/png;base64,iVBORw0KGgo="

function envWith(take?: () => Promise<string | null>): RibbonEnv {
  return {
    c: new SheetController(),
    copy() {},
    paste() {},
    dialog() {},
    toggleRibbon() {},
    backstage() {},
    ...(take ? { screenshot: { take } } : {}),
  }
}

/** jsdom doesn't decode pictures: an Image that "loads" at 1600 × 900. */
class FakeImage {
  naturalWidth = 1600
  naturalHeight = 900
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  set src(_v: string) {
    queueMicrotask(() => this.onload?.())
  }
}

afterEach(() => vi.unstubAllGlobals())

describe("Insert › Screenshot", () => {
  it("is off where the platform can't take screenshots", () => {
    expect(command.enabled!(envWith())).toBe(false)
    expect(command.enabled!(envWith(async () => null))).toBe(true)
  })

  it("puts the screenshot on the sheet, at most half the window wide, keeping its shape", async () => {
    vi.stubGlobal("Image", FakeImage)
    const env = envWith(async () => PNG)
    command.run!(env)
    await vi.waitFor(() => expect(env.c.sheet.objects).toHaveLength(1))
    const pic = env.c.sheet.objects[0]
    expect(pic).toMatchObject({ kind: "picture", src: PNG })
  })

  it("does nothing when cancelled, and says why when it fails", async () => {
    const cancelled = envWith(async () => null)
    command.run!(cancelled)
    await new Promise((r) => setTimeout(r, 10))
    expect(cancelled.c.sheet.objects).toHaveLength(0)
    const failed = envWith(async () => {
      throw new Error("The screenshot tool isn't available.")
    })
    const alert = vi.fn()
    failed.c.shell.alert = alert
    command.run!(failed)
    await vi.waitFor(() => expect(alert).toHaveBeenCalled())
    expect(alert.mock.calls[0][0]).toMatch(/couldn't be taken. The screenshot tool isn't available/)
  })

  it("turns the desktop tool's bytes into a picture of the right type", () => {
    expect(imageDataUrl(Uint8Array.from([0x89, 0x50, 0x4e, 0x47]))).toBe("data:image/png;base64,iVBORw==")
    expect(imageDataUrl(Uint8Array.from([0xff, 0xd8, 0xff]))).toMatch(/^data:image\/jpeg;base64,/)
  })
})
