import type { IconSetName } from "@cells/core"
import type { CellPainter as Painter } from "./draw-cells"

/**
 * Icon set glyphs, drawn on the canvas at any size. Index 0 is the lowest band (red, down), the last the highest
 * (green, up), as Excel orders a set's icons from its thresholds.
 */
const GREEN = "#3A9F49"
const YELLOW = "#E9B127"
const RED = "#D8423C"
const GRAY = "#7F7F7F"
const GOLD = "#F2B705"

/** Colour per band: red, yellow, green for three; with yellows in between for four and five. */
function bandColor(index: number, count: number): string {
  if (index === count - 1) return GREEN
  if (index === 0) return RED
  return YELLOW
}

/** An arrow pointing at `angle` (0 = up, 90 = right, 180 = down), in a size×size box at (x, y). */
function arrow(p: Painter, x: number, y: number, size: number, angle: number, color: string) {
  p.save()
  p.translate(x + size / 2, y + size / 2)
  p.rotate((angle * Math.PI) / 180)
  const h = size * 0.46
  const w = size * 0.42
  p.fillStyle = color
  p.beginPath()
  p.moveTo(0, -h)
  p.lineTo(w, -h + w)
  p.lineTo(w * 0.42, -h + w)
  p.lineTo(w * 0.42, h)
  p.lineTo(-w * 0.42, h)
  p.lineTo(-w * 0.42, -h + w)
  p.lineTo(-w, -h + w)
  p.closePath()
  p.fill()
  p.restore()
}

function circle(p: Painter, cx: number, cy: number, r: number) {
  p.beginPath()
  const steps = 24
  for (let i = 0; i <= steps; i++) {
    const a = (i / steps) * Math.PI * 2
    if (i === 0) p.moveTo(cx + r, cy)
    else p.lineTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r)
  }
  p.closePath()
}

function star(p: Painter, cx: number, cy: number, r: number) {
  p.beginPath()
  for (let i = 0; i < 10; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / 5
    const rr = i % 2 ? r * 0.45 : r
    if (i === 0) p.moveTo(cx + Math.cos(a) * rr, cy + Math.sin(a) * rr)
    else p.lineTo(cx + Math.cos(a) * rr, cy + Math.sin(a) * rr)
  }
  p.closePath()
}

export function drawIcon(p: Painter, set: IconSetName, index: number, x: number, y: number, size: number) {
  const cx = x + size / 2
  const cy = y + size / 2
  const r = size * 0.42
  switch (set) {
    case "3Arrows":
    case "3ArrowsGray": {
      const angle = [180, 90, 0][index]
      arrow(p, x, y, size, angle, set === "3ArrowsGray" ? GRAY : bandColor(index, 3))
      return
    }
    case "4Arrows":
      arrow(p, x, y, size, [180, 135, 45, 0][index], bandColor(index, 4))
      return
    case "5Arrows":
      arrow(p, x, y, size, [180, 135, 90, 45, 0][index], bandColor(index, 5))
      return
    case "3TrafficLights1":
      p.fillStyle = bandColor(index, 3)
      circle(p, cx, cy, r)
      p.fill()
      return
    case "3Signs":
      p.fillStyle = bandColor(index, 3)
      p.beginPath()
      if (index === 0) {
        p.moveTo(cx, cy - r)
        p.lineTo(cx + r, cy)
        p.lineTo(cx, cy + r)
        p.lineTo(cx - r, cy)
      } else if (index === 1) {
        p.moveTo(cx, cy - r)
        p.lineTo(cx + r, cy + r * 0.8)
        p.lineTo(cx - r, cy + r * 0.8)
      } else {
        circle(p, cx, cy, r)
      }
      p.closePath()
      p.fill()
      return
    case "3Symbols": {
      p.fillStyle = bandColor(index, 3)
      circle(p, cx, cy, r)
      p.fill()
      p.strokeStyle = "#ffffff"
      p.lineWidth = Math.max(1.2, size / 9)
      p.beginPath()
      const k = r * 0.45
      if (index === 0) {
        p.moveTo(cx - k, cy - k)
        p.lineTo(cx + k, cy + k)
        p.moveTo(cx + k, cy - k)
        p.lineTo(cx - k, cy + k)
      } else if (index === 1) {
        p.moveTo(cx, cy - k * 1.2)
        p.lineTo(cx, cy + k * 0.3)
        p.moveTo(cx, cy + k * 0.8)
        p.lineTo(cx, cy + k * 1.1)
      } else {
        p.moveTo(cx - k, cy)
        p.lineTo(cx - k * 0.2, cy + k * 0.8)
        p.lineTo(cx + k, cy - k * 0.7)
      }
      p.stroke()
      return
    }
    case "3Flags": {
      p.fillStyle = bandColor(index, 3)
      p.beginPath()
      p.moveTo(x + size * 0.28, y + size * 0.12)
      p.lineTo(x + size * 0.86, y + size * 0.32)
      p.lineTo(x + size * 0.28, y + size * 0.55)
      p.closePath()
      p.fill()
      p.fillStyle = GRAY
      p.fillRect(x + size * 0.2, y + size * 0.1, Math.max(1, size * 0.09), size * 0.8)
      return
    }
    case "3Stars": {
      star(p, cx, cy, r)
      p.strokeStyle = GOLD
      p.lineWidth = 1
      p.stroke()
      if (index === 0) return
      p.save()
      if (index === 1) {
        p.beginPath()
        p.rect(x, y, size / 2, size)
        p.clip()
      }
      star(p, cx, cy, r)
      p.fillStyle = GOLD
      p.fill()
      p.restore()
      return
    }
    case "5Quarters": {
      p.strokeStyle = GRAY
      p.lineWidth = 1
      circle(p, cx, cy, r)
      p.stroke()
      if (index === 0) return
      p.fillStyle = GRAY
      if (index === 4) {
        circle(p, cx, cy, r)
        p.fill()
        return
      }
      p.beginPath()
      p.moveTo(cx, cy)
      const end = -Math.PI / 2 + (index / 4) * Math.PI * 2
      const steps = 16
      for (let i = 0; i <= steps; i++) {
        const a = -Math.PI / 2 + ((end + Math.PI / 2) * i) / steps
        p.lineTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r)
      }
      p.closePath()
      p.fill()
      return
    }
  }
}

/** Each set's icons by name, lowest band first (for lists that can't draw them, such as the Sort dialog's). */
const ICON_NAMES: Record<IconSetName, string[]> = {
  "3Arrows": ["Red Down Arrow", "Yellow Side Arrow", "Green Up Arrow"],
  "3ArrowsGray": ["Gray Down Arrow", "Gray Side Arrow", "Gray Up Arrow"],
  "4Arrows": ["Red Down Arrow", "Yellow Down-Right Arrow", "Yellow Up-Right Arrow", "Green Up Arrow"],
  "5Arrows": [
    "Red Down Arrow",
    "Yellow Down-Right Arrow",
    "Yellow Side Arrow",
    "Yellow Up-Right Arrow",
    "Green Up Arrow",
  ],
  "3TrafficLights1": ["Red Circle", "Yellow Circle", "Green Circle"],
  "3Signs": ["Red Diamond", "Yellow Triangle", "Green Circle"],
  "3Symbols": ["Red Cross Symbol", "Yellow Exclamation Symbol", "Green Check Symbol"],
  "3Flags": ["Red Flag", "Yellow Flag", "Green Flag"],
  "3Stars": ["Silver Star", "Half Gold Star", "Gold Star"],
  "5Quarters": [
    "White Circle (All White Quarters)",
    "Circle with One Black Quarter",
    "Circle with Two Black Quarters",
    "Circle with Three Black Quarters",
    "Black Circle",
  ],
}

export function iconName(icon: { set: IconSetName; index: number } | null): string {
  return icon ? (ICON_NAMES[icon.set]?.[icon.index] ?? `Icon ${icon.index + 1}`) : "No Cell Icon"
}
