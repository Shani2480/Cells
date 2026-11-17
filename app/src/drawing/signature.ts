import type { SignatureSetup, Theme } from "@cells/core"

/** Insert › Signature Line's size, as Office makes it (2 × 1 inch). */
export const SIGNATURE_SIZE = { w: 192, h: 96 }

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
const n = (v: number) => String(Math.round(v * 100) / 100)

/**
 * An unsigned signature line, as Office draws it: a large X over the line at the left, and under the line the
 * suggested signer and their title. Drawn in its box (0,0)–(w,h) in the theme's body font.
 */
export function signatureSvg(setup: SignatureSetup, w: number, h: number, theme: Theme): string {
  const k = Math.min(w / SIGNATURE_SIZE.w, h / SIGNATURE_SIZE.h)
  const font = `font-family="${esc(theme.fonts.minor)}, Calibri, Carlito, sans-serif"`
  const lineY = h * 0.56
  const pad = 6 * k
  const x = `<text x="${n(pad)}" y="${n(lineY - 6 * k)}" ${font} font-size="${n(30 * k)}" fill="#000000">X</text>`
  const line = `<line x1="${n(pad)}" y1="${n(lineY)}" x2="${n(w - pad)}" y2="${n(lineY)}" stroke="#000000" stroke-width="${n(Math.max(0.75, 1.25 * k))}"/>`
  const text = (s: string, row: number) =>
    s
      ? `<text x="${n(pad)}" y="${n(lineY + (14 + row * 14) * k)}" ${font} font-size="${n(11 * k)}" fill="#000000">${esc(s)}</text>`
      : ""
  return `<rect x="0" y="0" width="${n(w)}" height="${n(h)}" fill="#FFFFFF"/>${x}${line}${text(setup.signer, 0)}${text(setup.title, 1)}`
}

/** A whole SVG document of the line, for turning into the PNG files carry. */
export function signatureDocument(setup: SignatureSetup, w: number, h: number, theme: Theme): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${n(w)}" height="${n(h)}" viewBox="0 0 ${n(w)} ${n(h)}">${signatureSvg(setup, w, h, theme)}</svg>`
}

/** A new line's GUID. */
export function signatureId(): string {
  const hex = [...crypto.getRandomValues(new Uint8Array(16))].map((b) => b.toString(16).padStart(2, "0")).join("")
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`.toUpperCase()
}
