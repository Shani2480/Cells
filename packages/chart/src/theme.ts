// From Inkspire (src/editor/shape-style.ts): the theme accents and shading charts use.
export const THEME_ACCENTS = [
  { id: "dark", label: "Black, Dark 1", color: "#000000" },
  { id: "accent1", label: "Blue, Accent 1", color: "#4472c4" },
  { id: "accent2", label: "Orange, Accent 2", color: "#ed7d31" },
  { id: "accent3", label: "Gray, Accent 3", color: "#a5a5a5" },
  { id: "accent4", label: "Gold, Accent 4", color: "#ffc000" },
  { id: "accent5", label: "Blue, Accent 5", color: "#5b9bd5" },
  { id: "accent6", label: "Green, Accent 6", color: "#70ad47" },
]

export function shade(color: string, amount: number) {
  const match = /^#?([0-9a-f]{6})$/i.exec(color.trim())
  if (!match) {
    return color
  }
  const value = Number.parseInt(match[1]!, 16)
  const channel = (shift: number) => {
    const part = (value >> shift) & 255
    const next = amount >= 0 ? part + (255 - part) * amount : part * (1 + amount)
    return Math.round(Math.min(255, Math.max(0, next)))
      .toString(16)
      .padStart(2, "0")
  }
  return `#${channel(16)}${channel(8)}${channel(0)}`
}
