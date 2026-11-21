/**
 * Insert › Pictures: a picture file picked on this device, read as a data URL with its size in pixels. Works the
 * same in the desktop app, on Android and in a browser (the file picker of the web view).
 */
export type PickedPicture = { src: string; w: number; h: number; name: string }

export function pickPicture(): Promise<PickedPicture | null> {
  return new Promise((resolve) => {
    const input = document.createElement("input")
    input.type = "file"
    input.accept = "image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml"
    input.onchange = () => {
      const file = input.files?.[0]
      if (!file) return resolve(null)
      void readPicture(file).then(resolve, () => resolve(null))
    }
    input.oncancel = () => resolve(null)
    input.click()
  })
}

export function readPicture(file: Blob & { name?: string }): Promise<PickedPicture> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(reader.error)
    reader.onload = () => {
      const src = String(reader.result)
      const img = new Image()
      img.onload = () => resolve({ src, w: img.naturalWidth || 96, h: img.naturalHeight || 96, name: file.name ?? "" })
      img.onerror = () => reject(new Error("That file isn't a picture Cells can show."))
      img.src = src
    }
    reader.readAsDataURL(file)
  })
}

/** A picture's size in pixels, from its data URL. */
export function pictureSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ w: img.naturalWidth || 96, h: img.naturalHeight || 96 })
    img.onerror = () => reject(new Error("The screenshot isn't a picture Cells can show."))
    img.src = src
  })
}

/** An SVG drawn into a PNG at a size (the copy older apps show for an SVG picture). */
export function svgToPng(svg: string, w: number, h: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    // A webview that never answers mustn't leave the caller waiting: the SVG stands in after a moment.
    const timer = setTimeout(() => reject(new Error("The picture took too long to draw.")), 3000)
    img.onload = () => {
      clearTimeout(timer)
      const canvas = document.createElement("canvas")
      canvas.width = Math.max(1, Math.round(w))
      canvas.height = Math.max(1, Math.round(h))
      const ctx = canvas.getContext("2d")
      if (!ctx) return reject(new Error("No canvas to draw the picture on."))
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
      resolve(canvas.toDataURL("image/png"))
    }
    img.onerror = () => {
      clearTimeout(timer)
      reject(new Error("The picture couldn't be drawn."))
    }
    img.src = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`
  })
}
