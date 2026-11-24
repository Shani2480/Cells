/**
 * Picture Format › Corrections and Color, Office 2010's kind: the picture's pixels changed as Excel changes them, for
 * Cells to show and to save (Office keeps the corrected picture beside the original). Measured against Excel 365:
 * brightness, contrast and colour tone are Excel's own curves (picture-curves.ts); saturation is HSL saturation scaled
 * (above 100%, raised to 1/s), which gives Excel's colours exactly; sharpen and soften are close to Excel's, not exact.
 *
 * Office applies them in the order the file lists them; Cells writes and applies sharpen or soften, colour tone,
 * saturation, then brightness and contrast.
 */
import { correctionsOf, type PictureAdjust } from "@cells/core"
import { BRIGHTNESS_CONTRAST, TEMPERATURES, TEMPERATURE_CURVES } from "./picture-curves"
import { applyArtistic } from "./picture-artistic"

export type Curves = { bc: Uint8Array; temps: Uint8Array }

async function inflate(base64: string): Promise<Uint8Array> {
  const bin = atob(base64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  const stream = new ReadableStream<BufferSource>({
    start(controller) {
      controller.enqueue(bytes)
      controller.close()
    },
  }).pipeThrough(new DecompressionStream("deflate"))
  const out = new Uint8Array(await new Response(stream).arrayBuffer())
  // Stored as differences, curve by curve.
  for (let i = 0; i < out.length; i++) if (i % 256) out[i] = (out[i] + out[i - 1]) & 255
  return out
}

let curves: Promise<Curves> | null = null
/** Excel's curves, unpacked once. */
export function loadCurves(): Promise<Curves> {
  curves ??= Promise.all([inflate(BRIGHTNESS_CONTRAST), inflate(TEMPERATURE_CURVES)]).then(([bc, temps]) => ({
    bc,
    temps,
  }))
  return curves
}

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v))

/** The curve for a brightness and contrast (each -1…1), between the measured ones. */
export function brightnessContrastCurve(c: Curves, bright: number, contrast: number): Uint8Array {
  const bi = clamp((bright + 1) * 10, 0, 20)
  const ci = clamp((contrast + 1) * 10, 0, 20)
  const b0 = Math.floor(bi)
  const c0 = Math.floor(ci)
  const b1 = Math.min(20, b0 + 1)
  const c1 = Math.min(20, c0 + 1)
  const fb = bi - b0
  const fc = ci - c0
  const at = (b: number, k: number, x: number) => c.bc[(b * 21 + k) * 256 + x]
  const out = new Uint8Array(256)
  for (let x = 0; x < 256; x++) {
    const top = at(b0, c0, x) * (1 - fc) + at(b0, c1, x) * fc
    const bottom = at(b1, c0, x) * (1 - fc) + at(b1, c1, x) * fc
    out[x] = Math.round(top * (1 - fb) + bottom * fb)
  }
  return out
}

/** The red, green and blue curves for a colour tone (kelvin), between the measured ones; 6500 changes nothing. */
export function temperatureCurves(c: Curves, kelvin: number): [Uint8Array, Uint8Array, Uint8Array] {
  // The measured temperatures, with 6500 (the picture as it is) in its place.
  const list = [...TEMPERATURES, 6500].sort((a, b) => a - b)
  const curve = (t: number, ch: number, x: number) => {
    if (t === 6500) return x
    return c.temps[(TEMPERATURES.indexOf(t) * 3 + ch) * 256 + x]
  }
  const k = clamp(kelvin, list[0], list[list.length - 1])
  let i = 0
  while (i < list.length - 2 && list[i + 1] < k) i++
  const t0 = list[i]
  const t1 = list[i + 1]
  const f = t1 === t0 ? 0 : (k - t0) / (t1 - t0)
  const out = [new Uint8Array(256), new Uint8Array(256), new Uint8Array(256)] as [Uint8Array, Uint8Array, Uint8Array]
  for (let ch = 0; ch < 3; ch++)
    for (let x = 0; x < 256; x++) out[ch][x] = Math.round(curve(t0, ch, x) * (1 - f) + curve(t1, ch, x) * f)
  return out
}

/** Points along which Excel's sharpen and soften grow (strength 0–1 → how much of the edge difference is added). */
const SHARPEN: [number, number][] = [
  [0, 0],
  [0.25, 0],
  [0.5, 0.125],
  [0.75, 1.29],
  [1, 2.58],
]
const SOFTEN: [number, number][] = [
  [0, 0],
  [0.25, 0],
  [0.5, 0.21],
  [0.75, 0.58],
  [1, 0.875],
]
function along(points: [number, number][], v: number): number {
  for (let i = 1; i < points.length; i++) {
    const [x0, y0] = points[i - 1]
    const [x1, y1] = points[i]
    if (v <= x1) return y0 + ((y1 - y0) * (v - x0)) / (x1 - x0)
  }
  return points[points.length - 1][1]
}

/** Sharpen (above 0) or soften (below): each pixel moved from, or toward, the mean of its four neighbours. */
export function sharpenSoften(data: Uint8ClampedArray, w: number, h: number, amount: number) {
  const k = amount > 0 ? along(SHARPEN, amount) : -along(SOFTEN, -amount)
  if (!k) return
  const src = data.slice()
  for (let y = 0; y < h; y++) {
    const up = (y > 0 ? y - 1 : y) * w
    const down = (y < h - 1 ? y + 1 : y) * w
    for (let x = 0; x < w; x++) {
      const left = x > 0 ? x - 1 : x
      const right = x < w - 1 ? x + 1 : x
      const i = (y * w + x) * 4
      for (let ch = 0; ch < 3; ch++) {
        const v = src[i + ch]
        const mean =
          (src[(up + x) * 4 + ch] +
            src[(down + x) * 4 + ch] +
            src[(y * w + left) * 4 + ch] +
            src[(y * w + right) * 4 + ch]) /
          4
        data[i + ch] = v + (v - mean) * k
      }
    }
  }
}

/** Colour saturation: HSL saturation times `s` (up to 1), or raised to 1/s (above): Excel's exactly. */
export function saturate(data: Uint8ClampedArray, s: number) {
  for (let i = 0; i < data.length; i += 4) {
    const r = data[i] / 255
    const g = data[i + 1] / 255
    const b = data[i + 2] / 255
    const max = Math.max(r, g, b)
    const min = Math.min(r, g, b)
    const d = max - min
    if (!d) continue
    const l = (max + min) / 2
    const sat = d / (1 - Math.abs(2 * l - 1))
    const k = s <= 1 ? s : Math.min(1, sat) ** (1 / s) / sat
    data[i] = (l + (r - l) * k) * 255 + 0.5
    data[i + 1] = (l + (g - l) * k) * 255 + 0.5
    data[i + 2] = (l + (b - l) * k) * 255 + 0.5
  }
}

function applyCurves(data: Uint8ClampedArray, curves: [Uint8Array, Uint8Array, Uint8Array]) {
  for (let i = 0; i < data.length; i += 4) {
    data[i] = curves[0][data[i]]
    data[i + 1] = curves[1][data[i + 1]]
    data[i + 2] = curves[2][data[i + 2]]
  }
}

/**
 * A picture's pixels (RGBA) with its corrections applied, in Office's order: the artistic effect first, unless the
 * file had it after the corrections. `unit`: pixels per Office pixel, for the artistic effect.
 */
export function correctPixels(data: Uint8ClampedArray, w: number, h: number, a: PictureAdjust, c: Curves, unit = 1) {
  if (a.artistic && !a.artistic.last) applyArtistic(data, w, h, a.artistic, unit)
  if (a.sharpen) sharpenSoften(data, w, h, a.sharpen)
  if (a.temperature !== undefined && a.temperature !== 6500) applyCurves(data, temperatureCurves(c, a.temperature))
  if (a.saturation !== undefined && a.saturation !== 1) saturate(data, a.saturation)
  if (a.bright || a.contrast) {
    const curve = brightnessContrastCurve(c, a.bright ?? 0, a.contrast ?? 0)
    applyCurves(data, [curve, curve, curve])
  }
  if (a.artistic?.last) applyArtistic(data, w, h, a.artistic, unit)
}

export { correctionsOf }

/** How a picture is corrected: at most `maxSide` pixels across, `unit` pixels for each of Office's. */
export type CorrectOptions = { maxSide?: number; unit?: number }

/**
 * A picture (data URL) with corrections applied, as a PNG data URL, in the browser (a canvas decodes and encodes it).
 */
export async function correctPicture(src: string, a: PictureAdjust, options: CorrectOptions = {}): Promise<string> {
  const c = await loadCurves()
  const image = await loadImage(src)
  const scale = Math.min(1, (options.maxSide ?? Infinity) / Math.max(image.naturalWidth, image.naturalHeight))
  const w = Math.max(1, Math.round(image.naturalWidth * scale))
  const h = Math.max(1, Math.round(image.naturalHeight * scale))
  const canvas = document.createElement("canvas")
  canvas.width = w
  canvas.height = h
  const ctx = canvas.getContext("2d", { willReadFrequently: true })!
  ctx.drawImage(image, 0, 0, w, h)
  const pixels = ctx.getImageData(0, 0, w, h)
  correctPixels(pixels.data, w, h, a, c, options.unit)
  ctx.putImageData(pixels, 0, 0)
  return canvas.toDataURL("image/png")
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error("The picture can't be read"))
    image.src = src
  })
}

/** The colour of a picture's pixel at (fx, fy), fractions of its width and height ("RRGGBB"). */
export async function pixelColor(src: string, fx: number, fy: number): Promise<string> {
  const image = await loadImage(src)
  const canvas = document.createElement("canvas")
  canvas.width = 1
  canvas.height = 1
  const ctx = canvas.getContext("2d", { willReadFrequently: true })!
  const x = clamp(Math.floor(fx * image.naturalWidth), 0, image.naturalWidth - 1)
  const y = clamp(Math.floor(fy * image.naturalHeight), 0, image.naturalHeight - 1)
  ctx.drawImage(image, x, y, 1, 1, 0, 0, 1, 1)
  const [r, g, b] = ctx.getImageData(0, 0, 1, 1).data
  return [r, g, b]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()
}
