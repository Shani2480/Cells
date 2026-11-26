/**
 * Reading a picture from the clipboard (Data › From Picture › Picture From Clipboard): Tauri's clipboard plugin on the
 * desktop (it hands over the pixels), the browser's clipboard in a browser. Android's clipboard is text only here.
 */

async function pngOf(rgba: Uint8Array, width: number, height: number): Promise<Blob | null> {
  const canvas = document.createElement("canvas")
  canvas.width = width
  canvas.height = height
  const ctx = canvas.getContext("2d")
  if (!ctx) return null
  ctx.putImageData(new ImageData(new Uint8ClampedArray(rgba), width, height), 0, 0)
  return new Promise((resolve) => canvas.toBlob(resolve, "image/png"))
}

export function clipboardImageReader(tauri: boolean, os: string): (() => Promise<Blob | null>) | null {
  if (tauri) {
    if (os === "android") return null
    return async () => {
      const { readImage } = await import("@tauri-apps/plugin-clipboard-manager")
      try {
        const image = await readImage()
        const { width, height } = await image.size()
        return await pngOf(await image.rgba(), width, height)
      } catch {
        // No picture on the clipboard (only text, or nothing).
        return null
      }
    }
  }
  if (typeof navigator === "undefined" || !navigator.clipboard?.read) return null
  return async () => {
    for (const item of await navigator.clipboard.read()) {
      const type = item.types.find((t) => t.startsWith("image/"))
      if (type) return item.getType(type)
    }
    return null
  }
}
