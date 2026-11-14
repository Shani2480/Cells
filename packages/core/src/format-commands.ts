import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { Op } from "./ops"
import type { Border, StylePatch } from "./styles"

/**
 * What Home-tab formatting commands do, as Excel does them: border presets, Increase/Decrease Decimal, Grow/Shrink
 * Font. Each returns operations or a patch; the controller applies them to the selection.
 */

export type BorderPreset =
  | "bottom"
  | "top"
  | "left"
  | "right"
  | "none"
  | "all"
  | "outside"
  | "thickOutside"
  | "bottomDouble"
  | "thickBottom"
  | "topBottom"
  | "topThickBottom"
  | "topDoubleBottom"
  | "insideHorizontal"
  | "insideVertical"

const THIN: Border = { style: "thin" }

const rowOf = (r: Range, row: number): Range => ({ r0: row, r1: row, c0: r.c0, c1: r.c1 })
const colOf = (r: Range, col: number): Range => ({ r0: r.r0, r1: r.r1, c0: col, c1: col })

/**
 * The operations for a Borders-menu preset over `range`. `line` is the pen chosen under Line Color / Line Style (a
 * thin automatic line by default); presets that name a weight (Thick, Double) use it with the pen's colour.
 */
export function borderOps(sheet: number, range: Range, preset: BorderPreset, line: Border = THIN): Op[] {
  const ops: Op[] = []
  const format = (ranges: Range[], patch: StylePatch) => ops.push({ kind: "format", sheet, ranges, patch })
  const thick: Border = { ...line, style: "thick" }
  const double: Border = { ...line, style: "double" }
  const top = rowOf(range, range.r0)
  const bottom = rowOf(range, range.r1)
  const left = colOf(range, range.c0)
  const right = colOf(range, range.c1)
  switch (preset) {
    case "bottom":
      format([bottom], { borderBottom: line })
      break
    case "top":
      format([top], { borderTop: line })
      break
    case "left":
      format([left], { borderLeft: line })
      break
    case "right":
      format([right], { borderRight: line })
      break
    case "none":
      format([range], {
        borderTop: null,
        borderBottom: null,
        borderLeft: null,
        borderRight: null,
        borderDiagonal: null,
        diagonalUp: null,
        diagonalDown: null,
      })
      // The neighbours' edges that face the range go too, so no line is left around it.
      if (range.r0 > 0) format([rowOf(range, range.r0 - 1)], { borderBottom: null })
      if (range.r1 < MAX_ROWS - 1) format([rowOf(range, range.r1 + 1)], { borderTop: null })
      if (range.c0 > 0) format([colOf(range, range.c0 - 1)], { borderRight: null })
      if (range.c1 < MAX_COLS - 1) format([colOf(range, range.c1 + 1)], { borderLeft: null })
      break
    case "all":
      format([range], { borderTop: line, borderBottom: line, borderLeft: line, borderRight: line })
      break
    case "outside":
    case "thickOutside": {
      const pen = preset === "outside" ? line : thick
      format([top], { borderTop: pen })
      format([bottom], { borderBottom: pen })
      format([left], { borderLeft: pen })
      format([right], { borderRight: pen })
      break
    }
    case "bottomDouble":
      format([bottom], { borderBottom: double })
      break
    case "thickBottom":
      format([bottom], { borderBottom: thick })
      break
    case "topBottom":
      format([top], { borderTop: line })
      format([bottom], { borderBottom: line })
      break
    case "topThickBottom":
      format([top], { borderTop: line })
      format([bottom], { borderBottom: thick })
      break
    case "topDoubleBottom":
      format([top], { borderTop: line })
      format([bottom], { borderBottom: double })
      break
    case "insideHorizontal":
      if (range.r1 > range.r0) format([{ ...range, r1: range.r1 - 1 }], { borderBottom: line })
      break
    case "insideVertical":
      if (range.c1 > range.c0) format([{ ...range, c1: range.c1 - 1 }], { borderRight: line })
      break
  }
  return ops
}

/** Excel's font size list (Grow Font / Shrink Font step through it). */
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

/** The next size up or down the list from `size` (past the ends, Excel steps by 10). */
export function stepFontSize(size: number, dir: 1 | -1): number {
  if (dir === 1) {
    const next = FONT_SIZES.find((s) => s > size)
    return next ?? Math.min(409, Math.floor(size / 10) * 10 + 10)
  }
  for (let i = FONT_SIZES.length - 1; i >= 0; i--) if (FONT_SIZES[i] < size) return FONT_SIZES[i]
  return Math.max(1, size - 1)
}

/**
 * Increase/Decrease Decimal. On General, Excel starts from what the cell shows (`shown`: 3.14159 shows five decimals,
 * so one more is "0.000000"; in a narrow column 1234567.891 shows "1234567.9", so one more is "0.00"). On a format, it
 * adds or removes one "0" after the decimal point in each number section; date, time, fraction and text sections stay
 * as they are. Returns null when nothing changes (checked against Excel with its own ribbon commands).
 */
export function changeDecimals(code: string | undefined, delta: 1 | -1, shown = ""): string | null {
  if (!code || code === "General") {
    const decimals = /E/i.test(shown) ? 2 : (shown.split(".")[1]?.length ?? 0)
    const next = decimals + delta
    if (next < 0) return null
    return next === 0 ? "0" : `0.${"0".repeat(next)}`
  }
  let changed = false
  const sections = splitSections(code).map((section) => {
    // Leave date/time and text sections alone.
    const bare = stripLiterals(section)
    if (bare.includes("/") || (/[dmyhs]/i.test(bare) && !/[0#?]/.test(bare))) return section
    const out = editSection(section, delta)
    if (out !== section) changed = true
    return out
  })
  return changed ? sections.join(";") : null
}

function stripLiterals(section: string) {
  return section.replace(/"[^"]*"|\\.|\[[^\]]*\]|_.|\*./g, "")
}

function splitSections(code: string): string[] {
  const out: string[] = []
  let cur = ""
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      const j = end < 0 ? code.length : end
      cur += code.slice(i, j + 1)
      i = j
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      cur += code.slice(i, i + 2)
      i++
    } else if (ch === ";") {
      out.push(cur)
      cur = ""
    } else cur += ch
  }
  out.push(cur)
  return out
}

/** Adds or removes one decimal place in one section, after the last digit placeholder of its number. */
function editSection(section: string, delta: 1 | -1): string {
  // Find the digit run that forms the number: the last placeholder outside quotes and brackets.
  const positions: number[] = []
  let point = -1
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      i = end < 0 ? section.length : end
    } else if (ch === "[") {
      const end = section.indexOf("]", i)
      i = end < 0 ? section.length : end
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      i++
    } else if (ch === "0" || ch === "#" || ch === "?") {
      positions.push(i)
    } else if (ch === "." && point < 0 && positions.length) {
      point = i
    } else if ((ch === "E" || ch === "e") && /[+-]/.test(section[i + 1] ?? "")) {
      break // exponent digits are not decimals
    }
  }
  if (!positions.length) return section
  if (delta === 1) {
    if (point < 0) {
      const at = positions[positions.length - 1] + 1
      return section.slice(0, at) + ".0" + section.slice(at)
    }
    // After the last decimal placeholder.
    const decimals = positions.filter((p) => p > point)
    const at = (decimals.length ? decimals[decimals.length - 1] : point) + 1
    return section.slice(0, at) + "0" + section.slice(at)
  }
  if (point < 0) return section
  const decimals = positions.filter((p) => p > point)
  if (!decimals.length) return section
  const last = decimals[decimals.length - 1]
  if (decimals.length === 1) return section.slice(0, point) + section.slice(last + 1)
  return section.slice(0, last) + section.slice(last + 1)
}
