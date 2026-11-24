/**
 * Picture Format › Artistic Effects: look-alikes of Office's 22 effects (Office's own filters aren't published, so
 * these are drawn to resemble Excel 365's, not to match them pixel for pixel). Like Excel's, they work in the
 * picture's pixels (Excel first brings a large picture down to its size on the sheet at 220 dpi; `unit` scales them
 * for a smaller copy, such as a gallery preview). Each effect's transparency lets the picture show through it.
 */
import { ARTISTIC_EFFECTS, type ArtisticEffect, type PictureArtistic } from "@cells/core"

/** A picture as three channels of 0–1. */
export type Planes = { w: number; h: number; r: Float32Array; g: Float32Array; b: Float32Array }

const clamp01 = (v: number) => (v < 0 ? 0 : v > 1 ? 1 : v)
const smooth = (e0: number, e1: number, x: number) => {
  const t = clamp01((x - e0) / (e1 - e0))
  return t * t * (3 - 2 * t)
}

export function toPlanes(data: Uint8ClampedArray, w: number, h: number): Planes {
  const n = w * h
  const p = { w, h, r: new Float32Array(n), g: new Float32Array(n), b: new Float32Array(n) }
  for (let i = 0; i < n; i++) {
    p.r[i] = data[i * 4] / 255
    p.g[i] = data[i * 4 + 1] / 255
    p.b[i] = data[i * 4 + 2] / 255
  }
  return p
}

function fromPlanes(p: Planes, data: Uint8ClampedArray) {
  for (let i = 0; i < p.w * p.h; i++) {
    data[i * 4] = p.r[i] * 255
    data[i * 4 + 1] = p.g[i] * 255
    data[i * 4 + 2] = p.b[i] * 255
  }
}

const copy = (p: Planes): Planes => ({ w: p.w, h: p.h, r: p.r.slice(), g: p.g.slice(), b: p.b.slice() })
const channels = (p: Planes) => [p.r, p.g, p.b]

/** The same random numbers every time (an effect looks the same each time it's drawn). */
function random(seed: number) {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) >>> 0
    let t = a
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

function noise(w: number, h: number, seed: number): Float32Array {
  const next = random(seed)
  const out = new Float32Array(w * h)
  for (let i = 0; i < out.length; i++) out[i] = next()
  return out
}

/** Smooth noise: random values on a grid `cell` pixels apart, blended between. */
function valueNoise(w: number, h: number, cell: number, seed: number): Float32Array {
  const gw = Math.ceil(w / cell) + 2
  const gh = Math.ceil(h / cell) + 2
  const grid = noise(gw, gh, seed)
  const out = new Float32Array(w * h)
  for (let y = 0; y < h; y++) {
    const fy = y / cell
    const y0 = Math.floor(fy)
    const ty = smooth(0, 1, fy - y0)
    for (let x = 0; x < w; x++) {
      const fx = x / cell
      const x0 = Math.floor(fx)
      const tx = smooth(0, 1, fx - x0)
      const a = grid[y0 * gw + x0] * (1 - tx) + grid[y0 * gw + x0 + 1] * tx
      const b = grid[(y0 + 1) * gw + x0] * (1 - tx) + grid[(y0 + 1) * gw + x0 + 1] * tx
      out[y * w + x] = a * (1 - ty) + b * ty
    }
  }
  return out
}

function boxPass(src: Float32Array, dst: Float32Array, w: number, h: number, r: number, horizontal: boolean) {
  const [len, lines] = horizontal ? [w, h] : [h, w]
  const at = (line: number, i: number) => (horizontal ? line * w + i : i * w + line)
  const size = 2 * r + 1
  for (let line = 0; line < lines; line++) {
    let sum = 0
    for (let i = -r; i <= r; i++) sum += src[at(line, Math.min(len - 1, Math.max(0, i)))]
    for (let i = 0; i < len; i++) {
      dst[at(line, i)] = sum / size
      sum += src[at(line, Math.min(len - 1, i + r + 1))] - src[at(line, Math.max(0, i - r))]
    }
  }
}

/** A Gaussian-like blur (three box blurs), `sigma` in pixels. */
function blur(a: Float32Array, w: number, h: number, sigma: number): Float32Array {
  const r = Math.round(sigma * 0.85)
  if (r < 1) return a.slice()
  const t = new Float32Array(a.length)
  const out = a.slice()
  for (let k = 0; k < 3; k++) {
    boxPass(out, t, w, h, r, true)
    boxPass(t, out, w, h, r, false)
  }
  return out
}

const blurPlanes = (p: Planes, sigma: number): Planes => ({
  w: p.w,
  h: p.h,
  r: blur(p.r, p.w, p.h, sigma),
  g: blur(p.g, p.w, p.h, sigma),
  b: blur(p.b, p.w, p.h, sigma),
})

function luminance(p: Planes): Float32Array {
  const out = new Float32Array(p.w * p.h)
  for (let i = 0; i < out.length; i++) out[i] = 0.2126 * p.r[i] + 0.7152 * p.g[i] + 0.0722 * p.b[i]
  return out
}

/** How strongly the picture changes at each pixel (Sobel), and in which direction. */
function edges(l: Float32Array, w: number, h: number) {
  const mag = new Float32Array(w * h)
  const gx = new Float32Array(w * h)
  const gy = new Float32Array(w * h)
  const v = (x: number, y: number) => l[Math.min(h - 1, Math.max(0, y)) * w + Math.min(w - 1, Math.max(0, x))]
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) {
      const dx =
        v(x + 1, y - 1) + 2 * v(x + 1, y) + v(x + 1, y + 1) - v(x - 1, y - 1) - 2 * v(x - 1, y) - v(x - 1, y + 1)
      const dy =
        v(x - 1, y + 1) + 2 * v(x, y + 1) + v(x + 1, y + 1) - v(x - 1, y - 1) - 2 * v(x, y - 1) - v(x + 1, y - 1)
      const i = y * w + x
      gx[i] = dx / 4
      gy[i] = dy / 4
      mag[i] = Math.hypot(dx, dy) / 4
    }
  return { mag, gx, gy }
}

/** Noise drawn out along a direction (strokes), `length` pixels long. */
function strokes(w: number, h: number, angle: number, length: number, seed: number): Float32Array {
  const n = noise(w, h, seed)
  const out = new Float32Array(w * h)
  const dx = Math.cos(angle)
  const dy = Math.sin(angle)
  const steps = Math.max(1, Math.round(length))
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) {
      let sum = 0
      for (let s = -steps; s <= steps; s++) {
        const sx = Math.min(w - 1, Math.max(0, Math.round(x + dx * s)))
        const sy = Math.min(h - 1, Math.max(0, Math.round(y + dy * s)))
        sum += n[sy * w + sx]
      }
      // Averages crowd toward the middle: spread them out again.
      out[y * w + x] = clamp01(0.5 + (sum / (2 * steps + 1) - 0.5) * Math.sqrt(2 * steps + 1) * 0.6)
    }
  return out
}

/** Light falling on a height map from the top left: 1 flat, more or less on slopes. */
function emboss(height: Float32Array, w: number, h: number, depth: number): Float32Array {
  const out = new Float32Array(w * h)
  const v = (x: number, y: number) => height[Math.min(h - 1, Math.max(0, y)) * w + Math.min(w - 1, Math.max(0, x))]
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) out[y * w + x] = 1 + (v(x - 1, y - 1) - v(x + 1, y + 1)) * depth
  return out
}

function each(p: Planes, f: (v: number, i: number, ch: number) => number) {
  channels(p).forEach((c, ch) => {
    for (let i = 0; i < c.length; i++) c[i] = clamp01(f(c[i], i, ch))
  })
}

function saturate(p: Planes, k: number) {
  for (let i = 0; i < p.r.length; i++) {
    const l = 0.2126 * p.r[i] + 0.7152 * p.g[i] + 0.0722 * p.b[i]
    p.r[i] = clamp01(l + (p.r[i] - l) * k)
    p.g[i] = clamp01(l + (p.g[i] - l) * k)
    p.b[i] = clamp01(l + (p.b[i] - l) * k)
  }
}

const quantize = (v: number, levels: number) => Math.round(v * (levels - 1)) / (levels - 1)

function grey(p: Planes, v: Float32Array) {
  p.r.set(v)
  p.g.set(v)
  p.b.set(v)
}

/** Samples the picture at other places (`dx`, `dy` per pixel). */
function displace(p: Planes, dx: Float32Array, dy: Float32Array): Planes {
  const out = copy(p)
  for (let y = 0; y < p.h; y++)
    for (let x = 0; x < p.w; x++) {
      const i = y * p.w + x
      const sx = Math.min(p.w - 1, Math.max(0, Math.round(x + dx[i])))
      const sy = Math.min(p.h - 1, Math.max(0, Math.round(y + dy[i])))
      const j = sy * p.w + sx
      out.r[i] = p.r[j]
      out.g[i] = p.g[j]
      out.b[i] = p.b[j]
    }
  return out
}

/** Each effect: the picture, its setting (Office's scale) and the pixel unit → the picture with the effect. */
const EFFECTS: Record<ArtisticEffect, (p: Planes, amount: number, u: number) => Planes> = {
  Blur: (p, radius, u) => blurPlanes(p, radius * 0.6 * u),

  Cement: (p, spacing, u) => {
    const { w, h } = p
    // Cracks between cells (the edges of a Voronoi pattern), on a rough surface.
    const cell = (6 + spacing * 0.35) * u
    const gw = Math.ceil(w / cell) + 1
    const gh = Math.ceil(h / cell) + 1
    const next = random(7)
    const pts = Array.from({ length: gw * gh }, (_, k) => [
      ((k % gw) + next()) * cell,
      (Math.floor(k / gw) + next()) * cell,
    ])
    const height = new Float32Array(w * h)
    const rough = valueNoise(w, h, 2 * u, 11)
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const cx = Math.floor(x / cell)
        const cy = Math.floor(y / cell)
        let d1 = Infinity
        let d2 = Infinity
        for (let j = cy - 1; j <= cy + 1; j++)
          for (let i = cx - 1; i <= cx + 1; i++) {
            if (i < 0 || j < 0 || i >= gw || j >= gh) continue
            const [px, py] = pts[j * gw + i]
            const d = Math.hypot(px - x, py - y)
            if (d < d1) [d1, d2] = [d, d1]
            else if (d < d2) d2 = d
          }
        height[y * w + x] = smooth(0, 1.6 * u, d2 - d1) * 0.4 + rough[y * w + x] * 0.6
      }
    const light = emboss(height, w, h, 2.6)
    const { mag } = edges(luminance(blurPlanes(p, u)), w, h)
    const out = copy(p)
    each(out, (v, i) => v * light[i] * (1 - Math.min(0.8, mag[i] * 3)))
    return out
  },

  CrisscrossEtching: (p, pressure, u) => {
    const out = blurPlanes(p, u)
    const l = luminance(out)
    const gap = 4 * u
    const strength = 0.1 + (pressure / 100) * 0.3
    each(out, (v, i) => {
      const x = i % p.w
      const y = Math.floor(i / p.w)
      const a = ((x + y) / gap) % 1 < 0.35 && l[i] < 0.8 ? 1 : 0
      const b = ((x - y + p.h * 4) / gap) % 1 < 0.35 && l[i] < 0.5 ? 1 : 0
      return (v * 0.7 + 0.3) * (1 - strength * Math.min(1, a + b))
    })
    return out
  },

  ChalkSketch: (p, pressure, u) => {
    const l = luminance(blurPlanes(p, 0.8 * u))
    const s = strokes(p.w, p.h, -Math.PI / 3, 8 * u, 3)
    const contrast = 1.3 + pressure * 0.3
    const out = copy(p)
    grey(
      out,
      l.map((v, i) => clamp01((v - 0.5) * contrast + 0.42 + (s[i] - 0.5) * 0.55)),
    )
    return out
  },

  Cutout: (p, shades, u) => {
    // A few flat colours: the picture in bands of brightness, each band its average colour.
    const soft = blurPlanes(p, 2 * u)
    const l = luminance(soft)
    const levels = shades + 2
    const sums = Array.from({ length: levels }, () => [0, 0, 0, 0])
    const band = l.map((v) => Math.min(levels - 1, Math.floor(v * levels)))
    band.forEach((k, i) => {
      sums[k][0] += soft.r[i]
      sums[k][1] += soft.g[i]
      sums[k][2] += soft.b[i]
      sums[k][3]++
    })
    const out = copy(p)
    each(out, (_, i, ch) => {
      const s = sums[band[i]]
      return s[ch] / Math.max(1, s[3])
    })
    return out
  },

  FilmGrain: (p, size, u) => {
    const out = blurPlanes(p, 0.6 * u)
    const amp = (size / 100) * 0.65
    const mono = noise(p.w, p.h, 5)
    channels(out).forEach((c, ch) => {
      const n = noise(p.w, p.h, 17 + ch)
      for (let i = 0; i < c.length; i++) c[i] = clamp01(c[i] + ((n[i] - 0.5) * 0.6 + (mono[i] - 0.5) * 0.6) * amp)
    })
    return out
  },

  Glass: (p, scaling, u) => {
    const cell = (3 + scaling * 0.12) * u
    const a = 5 * u
    const nx = valueNoise(p.w, p.h, cell, 21)
    const ny = valueNoise(p.w, p.h, cell, 22)
    return displace(
      p,
      nx.map((v) => (v - 0.5) * 2 * a),
      ny.map((v) => (v - 0.5) * 2 * a),
    )
  },

  GlowDiffused: (p, intensity, u) => {
    const out = blurPlanes(p, 1.2 * u)
    const l = luminance(out)
    const n = noise(p.w, p.h, 9)
    const k = intensity / 10
    each(out, (v, i) => {
      const glow = clamp01(k * (0.35 + 1.3 * l[i]) * (0.55 + 0.9 * n[i]))
      return v + (1 - v) * glow
    })
    return out
  },

  GlowEdges: (p, smoothness, u) => {
    const soft = blurPlanes(p, (0.6 + smoothness * 0.35) * u)
    const { mag } = edges(luminance(soft), p.w, p.h)
    const out = copy(soft)
    for (let i = 0; i < mag.length; i++) {
      const e = clamp01(mag[i] * 14)
      // The edge in the picture's own colour, more vivid and at full brightness.
      const l = 0.2126 * soft.r[i] + 0.7152 * soft.g[i] + 0.0722 * soft.b[i]
      const c = [soft.r[i], soft.g[i], soft.b[i]].map((v) => Math.max(0, l + (v - l) * 2.5))
      const top = Math.max(c[0], c[1], c[2], 0.05)
      out.r[i] = clamp01((c[0] / top) * e + soft.r[i] * 0.08 * (1 - e))
      out.g[i] = clamp01((c[1] / top) * e + soft.g[i] * 0.08 * (1 - e))
      out.b[i] = clamp01((c[2] / top) * e + soft.b[i] * 0.08 * (1 - e))
    }
    return out
  },

  LightScreen: (p, grid, u) => {
    const size = Math.max(2, Math.round((4 + grid) * u))
    const out = copy(p)
    const { w, h } = p
    for (let cy = 0; cy < h; cy += size)
      for (let cx = 0; cx < w; cx += size) {
        const avg = [0, 0, 0]
        let count = 0
        for (let y = cy; y < Math.min(h, cy + size); y++)
          for (let x = cx; x < Math.min(w, cx + size); x++) {
            const i = y * w + x
            avg[0] += p.r[i]
            avg[1] += p.g[i]
            avg[2] += p.b[i]
            count++
          }
        for (let y = cy; y < Math.min(h, cy + size); y++)
          for (let x = cx; x < Math.min(w, cx + size); x++) {
            const i = y * w + x
            // A bevelled tile: light at its top and left, dark at its bottom and right.
            const fx = (x - cx) / size
            const fy = (y - cy) / size
            const shade = fx < 0.15 || fy < 0.15 ? 1.25 : fx > 0.85 || fy > 0.85 ? 0.7 : 1.1
            channels(out).forEach((c, ch) => (c[i] = clamp01((avg[ch] / count) * shade * 0.7 + c[i] * 0.3 + 0.05)))
          }
      }
    return out
  },

  LineDrawing: (p, pencil, u) => {
    const { mag } = edges(luminance(blurPlanes(p, 0.7 * u)), p.w, p.h)
    const n = noise(p.w, p.h, 13)
    const grain = 0.15 + (pencil / 100) * 0.3
    const out = copy(p)
    each(out, (v, i, ch) => {
      const light = v * 0.85 + 0.15
      return light * (1 - clamp01(mag[i] * 4) * 0.4) + (n[i] - 0.5) * grain * (0.6 + ch * 0.2)
    })
    return out
  },

  Marker: (p, size, u) => {
    const out = blurPlanes(p, (0.5 + (size / 100) * 1.2) * u)
    const { mag } = edges(luminance(out), p.w, p.h)
    each(out, (v, i) => quantize(v, 14) * (1 - clamp01(mag[i] * 2) * 0.3))
    saturate(out, 1.15)
    return out
  },

  MosiaicBubbles: (p, pressure, u) => {
    const out = blurPlanes(p, 0.5 * u)
    const cell = 6 * u
    const next = random(31)
    const gw = Math.ceil(p.w / cell) + 1
    const jitter = Array.from({ length: gw * (Math.ceil(p.h / cell) + 1) }, () => [next() - 0.5, next() - 0.5])
    const k = 0.25 + (pressure / 100) * 0.5
    each(out, (v, i) => {
      const x = (i % p.w) / cell
      const y = Math.floor(i / p.w) / cell
      const [jx, jy] = jitter[Math.floor(y) * gw + Math.floor(x)]
      const dx = (x % 1) - 0.5 - jx * 0.3
      const dy = (y % 1) - 0.5 - jy * 0.3
      const d = Math.hypot(dx, dy) * 2
      const ring = smooth(0.6, 0.9, d) * (1 - smooth(0.9, 1.05, d))
      const shine = dx < 0 && dy < 0 ? ring * 0.5 : -ring * 0.4
      return v + shine * k
    })
    return out
  },

  PaintBrush: (p, brush, u) => {
    const out = blurPlanes(p, (0.6 + brush * 0.3) * u)
    const { mag } = edges(luminance(out), p.w, p.h)
    const s = strokes(p.w, p.h, Math.PI / 4, 3 * u, 41)
    each(out, (v, i) => quantize(v, 16) * (1 - clamp01(mag[i] * 4) * 0.6) + (s[i] - 0.5) * 0.03)
    return out
  },

  PaintStrokes: (p, intensity, u) => {
    const out = blurPlanes(p, 0.8 * u)
    const l = luminance(out)
    const s = strokes(p.w, p.h, -Math.PI / 4, 6 * u, 43)
    const k = intensity / 5
    each(out, (v, i) => {
      const keep = smooth(0.45, 0.72, l[i])
      return v * keep * 1.08 + (s[i] - 0.5) * 0.12 * k * keep
    })
    return out
  },

  PastelsSmooth: (p, scaling, u) => {
    const out = blurPlanes(p, (1.5 + scaling * 0.03) * u)
    saturate(out, 1.3)
    const s = strokes(p.w, p.h, Math.PI / 4, 7 * u, 47)
    const light = emboss(s, p.w, p.h, 0.8)
    each(out, (v, i) => v * light[i] + (s[i] - 0.5) * 0.08)
    return out
  },

  PencilGrayscale: (p, pencil, u) => {
    const l = luminance(blurPlanes(p, 0.6 * u))
    const { mag } = edges(l, p.w, p.h)
    const len = (3 + (pencil / 100) * 10) * u
    const a = strokes(p.w, p.h, Math.PI / 5, len, 51)
    const b = strokes(p.w, p.h, -Math.PI / 4, len, 52)
    const out = copy(p)
    grey(
      out,
      l.map((v, i) => {
        const dark = 1 - v
        const lines = Math.max(a[i], b[i] * (dark > 0.4 ? 1 : 0.5))
        return clamp01(0.92 - dark * 0.9 * (0.35 + lines) - (lines - 0.5) * 0.15 - clamp01(mag[i] * 4) * 0.45)
      }),
    )
    return out
  },

  PencilSketch: (p, pressure, u) => {
    const l = luminance(blurPlanes(p, 0.6 * u))
    const { mag } = edges(l, p.w, p.h)
    const n = noise(p.w, p.h, 61)
    const k = 0.6 + (pressure / 100) * 0.9
    const out = copy(p)
    grey(
      out,
      l.map((v, i) => {
        const ink = clamp01(((1 - v) * 1.1 + mag[i] * 5 - 0.38) * k)
        return n[i] < ink * 0.85 ? 0.3 + n[i] * 0.3 : 1
      }),
    )
    return blurPlanes(out, 0.4 * u)
  },

  Photocopy: (p, detail, u) => {
    const l = luminance(p)
    const near = blur(l, p.w, p.h, u)
    const far = blur(l, p.w, p.h, (2 + (10 - detail) * 0.6) * u)
    const out = copy(p)
    const tint = [1, 0.95, 0.88]
    each(out, (_, i, ch) => clamp01(0.82 + (near[i] - far[i]) * 10 + Math.min(0, l[i] - 0.35) * 2.2) * tint[ch])
    return out
  },

  PlasticWrap: (p, smoothness, u) => {
    const out = blurPlanes(p, u)
    const height = blur(luminance(p), p.w, p.h, (3 + smoothness * 0.8) * u)
    const { gx, gy } = edges(height, p.w, p.h)
    each(out, (v, i) => {
      const slope = -(gx[i] + gy[i]) * 18
      const shine = clamp01(slope - 0.25) ** 2
      const shadow = clamp01(-slope) * 0.35
      return v * 0.85 * (1 - shadow) + shine * 0.95
    })
    return out
  },

  Texturizer: (p, scaling, u) => {
    const period = (2 + scaling * 0.05) * u
    const rough = valueNoise(p.w, p.h, u, 71)
    const height = rough.map((v, i) => {
      const x = i % p.w
      const y = Math.floor(i / p.w)
      return v * 0.5 + 0.25 * (Math.sin((x / period) * Math.PI) * Math.sin((y / period) * Math.PI) + 1)
    })
    const light = emboss(height, p.w, p.h, 0.9)
    const out = copy(p)
    each(out, (v, i) => v * light[i])
    return out
  },

  WatercolorSponge: (p, brush, u) => {
    const soft = blurPlanes(p, (1 + brush * 0.4) * u)
    const nx = valueNoise(p.w, p.h, 5 * u, 81)
    const ny = valueNoise(p.w, p.h, 5 * u, 82)
    const out = displace(
      soft,
      nx.map((v) => (v - 0.5) * 4 * u),
      ny.map((v) => (v - 0.5) * 4 * u),
    )
    const { mag } = edges(luminance(out), p.w, p.h)
    const sponge = valueNoise(p.w, p.h, 3 * u, 83)
    const l = luminance(out)
    each(out, (v, i) => {
      // Blotches lighten and darken each pixel's colour as a whole, keeping its hue.
      const f = (1 + (sponge[i] - 0.5) * 0.35) * (1 - clamp01(mag[i] * 3) * 0.35)
      return v * f + (quantize(l[i], 10) - l[i]) * 0.5
    })
    saturate(out, 1.2)
    return out
  },
}

/** A picture's pixels (RGBA) with an artistic effect; `unit` is how many pixels stand for one of Office's. */
export function applyArtistic(data: Uint8ClampedArray, w: number, h: number, a: PictureArtistic, unit = 1) {
  const spec = ARTISTIC_EFFECTS[a.effect]
  if (!spec) return
  const before = toPlanes(data, w, h)
  const after = EFFECTS[a.effect](before, a.amount ?? spec.amount, Math.max(0.1, unit))
  const trans = spec.trans === null ? 0 : clamp01(a.trans ?? spec.trans)
  if (trans) {
    channels(after).forEach((c, ch) => {
      const o = channels(before)[ch]
      for (let i = 0; i < c.length; i++) c[i] = c[i] * (1 - trans) + o[i] * trans
    })
  }
  fromPlanes(after, data)
}
