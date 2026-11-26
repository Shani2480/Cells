import type { Range } from "./address"
import { shiftRange, subtract, ValueContext } from "./conditional"
import { evaluateFormula } from "./formula/evaluate"
import { parseFormula } from "./formula/parser"
import type { Node } from "./formula/ast"
import type { Scalar } from "./formula/types"
import { isError } from "./values"
import type { Sheet } from "./sheet"

/**
 * Hyperlinks, as Excel keeps them: a link over a range of cells (usually one) that goes to a web page, an e-mail
 * address, a file, or a place in this workbook. The cell's text is ordinary cell content; the link only says where
 * a click goes and what its ScreenTip says.
 */
export type Hyperlink = {
  range: Range
  /** Outside the workbook: a web address, a mailto: address, or a file (a full path, or relative to the workbook). */
  address?: string
  /** A place: "Sheet2!A1" or a defined name, in this workbook (or, with `address`, in that file). */
  location?: string
  /** The ScreenTip shown on hover instead of the address. */
  tooltip?: string
  /** Excel's display attribute: the text the link was made with. Kept for the file; the cell's text is what shows. */
  display?: string
}

/** What a link points at, for the dialog's "Link to" choice and for following it. */
export type HyperlinkKind = "web" | "mail" | "file" | "place"

export function copyHyperlinks(links: readonly Hyperlink[]): Hyperlink[] {
  return links.map((l) => ({ ...l, range: { ...l.range } }))
}

const inside = (r: Range, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1
const overlaps = (a: Range, b: Range) => a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/** The link on a cell (the last added wins, as in Excel's list), or null. */
export function hyperlinkAt(links: readonly Hyperlink[], row: number, col: number): Hyperlink | null {
  for (let i = links.length - 1; i >= 0; i--) if (inside(links[i].range, row, col)) return links[i]
  return null
}

/** The links that touch any of `ranges`. */
export function hyperlinksIn(links: readonly Hyperlink[], ranges: readonly Range[]): Hyperlink[] {
  return links.filter((l) => ranges.some((r) => overlaps(l.range, r)))
}

/**
 * The sheet's links after `ranges` get `link` (null: Remove). Cells in `ranges` leave the links they had; a link
 * over a larger range keeps the rest of its cells, split into rectangles as Excel does.
 */
export function withHyperlink(
  links: readonly Hyperlink[],
  ranges: readonly Range[],
  link: Omit<Hyperlink, "range"> | null,
): Hyperlink[] {
  const out: Hyperlink[] = []
  for (const l of links) {
    let parts = [{ ...l.range }]
    for (const cut of ranges) parts = parts.flatMap((p) => subtract(p, cut))
    for (const p of parts) out.push({ ...l, range: p })
  }
  if (link) for (const r of ranges) out.push({ ...clean(link), range: { ...r } })
  return out
}

/** Links follow inserted and deleted rows and columns; a link whose cells all go is dropped. */
export function shiftHyperlinks(links: readonly Hyperlink[], axis: "rows" | "cols", at: number, count: number) {
  const out: Hyperlink[] = []
  for (const l of links) {
    const range = shiftRange(l.range, axis, at, count)
    if (range) out.push({ ...l, range })
  }
  return out
}

/**
 * Insert or Delete Cells (count > 0 inserts, < 0 deletes at `at`), shifting only the `band` (first and last column
 * for rows, row for columns): a link wholly in the band moves with its cells (shrinks, or goes with deleted ones);
 * one that reaches outside the band stays where it is, as measured in Excel 365.
 */
export function shiftHyperlinksInBand(
  links: readonly Hyperlink[],
  axis: "rows" | "cols",
  at: number,
  count: number,
  band: [number, number],
): Hyperlink[] {
  const out: Hyperlink[] = []
  for (const l of links) {
    const [a0, a1] = axis === "rows" ? [l.range.c0, l.range.c1] : [l.range.r0, l.range.r1]
    if (a0 < band[0] || a1 > band[1]) {
      out.push(l)
      continue
    }
    const range = shiftRange(l.range, axis, at, count)
    if (range) out.push({ ...l, range })
  }
  return out
}

const within = (r: Range, outer: Range) => r.r0 >= outer.r0 && r.r1 <= outer.r1 && r.c0 >= outer.c0 && r.c1 <= outer.c1

/**
 * The links a copy (or, with `whole`, a cut) of `range` takes along, placed relative to its top-left cell: a copy
 * takes each link on its cells, trimmed to them; a cut only the links wholly inside (Excel leaves one that reaches
 * outside where it is).
 */
export function readHyperlinks(links: readonly Hyperlink[], range: Range, whole = false): Hyperlink[] {
  const out: Hyperlink[] = []
  for (const l of links) {
    if (!overlaps(l.range, range) || (whole && !within(l.range, range))) continue
    const r = {
      r0: Math.max(l.range.r0, range.r0) - range.r0,
      r1: Math.min(l.range.r1, range.r1) - range.r0,
      c0: Math.max(l.range.c0, range.c0) - range.c0,
      c1: Math.min(l.range.c1, range.c1) - range.c0,
    }
    out.push({ ...l, range: r })
  }
  return out
}

/**
 * The sheet's links after pasting a block of `rows` × `cols` carrying `block` (from readHyperlinks) over `area`,
 * repeated where the paste is tiled: the area's own links give way, as cells pasted in Excel replace theirs.
 */
export function pasteHyperlinks(
  links: readonly Hyperlink[],
  area: Range,
  block: readonly Hyperlink[],
  rows: number,
  cols: number,
): Hyperlink[] {
  const out = withHyperlink(links, [area], null)
  for (let r0 = area.r0; r0 <= area.r1; r0 += rows)
    for (let c0 = area.c0; c0 <= area.c1; c0 += cols)
      for (const l of block) {
        const range = {
          r0: r0 + l.range.r0,
          c0: c0 + l.range.c0,
          r1: Math.min(area.r1, r0 + l.range.r1),
          c1: Math.min(area.c1, c0 + l.range.c1),
        }
        if (range.r0 <= range.r1 && range.c0 <= range.c1) out.push({ ...l, range })
      }
  return out
}

/** The sheet's links once a cut of `range` has taken the ones wholly inside it away. */
export function withoutHyperlinksIn(links: readonly Hyperlink[], range: Range): Hyperlink[] {
  return links.filter((l) => !within(l.range, range))
}

/** Drops empty fields, so a link compares and saves the same however the dialog filled it. */
function clean(link: Omit<Hyperlink, "range">): Omit<Hyperlink, "range"> {
  const out: Omit<Hyperlink, "range"> = {}
  if (link.address?.trim()) out.address = link.address.trim()
  if (link.location?.trim()) out.location = link.location.trim().replace(/^#/, "")
  if (link.tooltip) out.tooltip = link.tooltip
  if (link.display) out.display = link.display
  return out
}

/** What kind of target the link has. */
export function hyperlinkKind(link: Pick<Hyperlink, "address" | "location">): HyperlinkKind {
  const address = link.address?.trim() ?? ""
  if (!address) return "place"
  if (/^mailto:/i.test(address)) return "mail"
  if (isWebAddress(address)) return "web"
  return "file"
}

/** http(s), ftp and other scheme://… addresses, and "www.…" that Excel turns into http://. */
export function isWebAddress(text: string): boolean {
  const t = text.trim()
  if (/^[a-z]:[\\/]/i.test(t)) return false
  return /^[a-z][a-z0-9+.-]*:\/\//i.test(t) || /^www\./i.test(t)
}

/**
 * What a typed address becomes, as Excel's dialog completes it: "www.example.com" → "http://www.example.com",
 * "name@example.com" → "mailto:name@example.com". Paths and full addresses stay as they are.
 */
export function completeAddress(text: string): string {
  const t = text.trim()
  if (/^www\./i.test(t)) return "http://" + t
  if (/^[^\s@/\\:]+@[^\s@/\\]+\.[^\s@/\\]+$/.test(t)) return "mailto:" + t
  return t
}

/** A mailto: address and its subject, for the E-mail Address page. */
export function mailParts(address: string): { to: string; subject: string } {
  const body = address.replace(/^mailto:/i, "")
  const q = body.indexOf("?")
  const to = decodeSafe(q < 0 ? body : body.slice(0, q))
  let subject = ""
  if (q >= 0) {
    for (const part of body.slice(q + 1).split("&")) {
      const eq = part.indexOf("=")
      if (eq > 0 && part.slice(0, eq).toLowerCase() === "subject") subject = decodeSafe(part.slice(eq + 1))
    }
  }
  return { to, subject }
}

/** The mailto: address Excel writes for an e-mail address and subject. */
export function mailAddress(to: string, subject: string): string {
  const t = to.trim().replace(/^mailto:/i, "")
  return "mailto:" + t + (subject.trim() ? "?subject=" + encodeURIComponent(subject.trim()) : "")
}

function decodeSafe(text: string): string {
  try {
    return decodeURIComponent(text)
  } catch {
    return text
  }
}

/**
 * The text a new link gives an empty cell, as Excel fills "Text to display": the address without mailto:, or the
 * place ("Sheet2!A1"), or the file and place ("Book.xlsx - Sheet2!A1").
 */
export function linkText(link: Pick<Hyperlink, "address" | "location">): string {
  const address = link.address?.trim() ?? ""
  const location = link.location?.trim().replace(/^#/, "") ?? ""
  if (!address) return location
  const shown = address.replace(/^mailto:/i, "")
  return location ? `${shown} - ${location}` : shown
}

/** The ScreenTip Excel shows when a link has none of its own: the address and place, then how to follow it. */
export function linkScreenTip(link: Pick<Hyperlink, "address" | "location" | "tooltip">, touch = false): string {
  if (link.tooltip) return link.tooltip
  const address = link.address?.trim() ?? ""
  const location = link.location?.trim() ?? ""
  const where = address ? address + (location ? " - " + location : "") : `#${location}`
  return touch ? where : `${where} - Click once to follow. Click and hold to select this cell.`
}

/**
 * HYPERLINK's link_location as a link: "#Sheet2!A1" is a place here, "[Book.xlsx]Sheet1!A1" and "Book.xlsx#Sheet1!A1"
 * a place in a file, anything else an address (a web address keeps its #fragment).
 */
export function linkFromText(text: string): Omit<Hyperlink, "range"> | null {
  const t = text.trim()
  if (!t) return null
  if (t.startsWith("#")) return t.length > 1 ? { location: t.slice(1) } : null
  const book = /^\[([^\]]+)\](.+)$/.exec(t)
  if (book) return { address: book[1], location: book[2] }
  if (!isWebAddress(t) && !/^mailto:/i.test(t)) {
    const hash = t.indexOf("#")
    if (hash > 0) return { address: t.slice(0, hash), location: t.slice(hash + 1) }
  }
  return { address: completeAddress(t) }
}

/**
 * The link of a HYPERLINK formula in a cell (=HYPERLINK("https://…", "Help"), its first argument worked out now), or
 * null when the cell's formula isn't one.
 */
export function formulaLink(
  sheets: readonly Sheet[],
  sheet: number,
  row: number,
  col: number,
  today: number,
): Omit<Hyperlink, "range"> | null {
  const input = sheets[sheet]?.cells.input(row, col)
  if (!input || !/HYPERLINK\s*\(/i.test(input)) return null
  const parsed = parseFormula(input.slice(1))
  if (!parsed.ok) return null
  let node: Node = parsed.ast
  while (node.type === "paren") node = node.arg
  if (node.type !== "func" || node.name.toUpperCase() !== "HYPERLINK" || !node.args.length) return null
  const ctx = new ValueContext(sheets, today)
  ctx.sheet = sheet
  ctx.row = row
  ctx.col = col
  const value = evaluateFormula(node.args[0], ctx) as Scalar
  if (value === null || isError(value) || typeof value === "object") return null
  return linkFromText(String(value))
}
