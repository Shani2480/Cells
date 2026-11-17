import {
  CLEAR_ALL,
  completeAddress,
  copyHyperlinks,
  formulaLink,
  hyperlinkAt,
  hyperlinkKind,
  hyperlinksIn,
  withHyperlink,
  type Hyperlink,
  type Op,
  type Range,
  type StylePatch,
} from "@cells/core"
import type { SheetController } from "./controller"
import { todaySerial } from "./controller"

/** What Insert › Link sets: where the link goes, its ScreenTip and (for a cell holding text or nothing) the text. */
export type LinkInit = {
  address?: string
  location?: string
  tooltip?: string
  /** "Text to display"; null when the cell's content can't be replaced (a number or a formula). */
  text: string | null
}

/** Where a click on a link goes, resolved for the shell: a place here, a workbook to open, or something for the system. */
export type LinkTarget =
  | { kind: "place"; location: string }
  | { kind: "workbook"; path: string; location?: string }
  | { kind: "external"; target: string; file: boolean }

/** Excel's Hyperlink cell style (Office 2023 theme): the theme's hyperlink colour, underlined. */
export const HYPERLINK_STYLE: StylePatch = { cellStyle: "Hyperlink", fontColor: "#467886", underline: "single" }

/** Excel's warning before a link opens a file. */
export const FILE_LINK_WARNING =
  "Hyperlinks can be harmful to your computer and data. To protect your computer, click only those hyperlinks from trusted sources. Do you want to continue?"

/** Excel's message when a link's place isn't in the workbook. */
export const BAD_REFERENCE = "Reference isn't valid."

/** Files a link opens in Cells itself rather than in another app. */
const WORKBOOK = /\.(xlsx|xlsm|xls|csv)$/i

/**
 * Insert › Link (Ctrl+K), the cell menu's Link items, Clear and Remove Hyperlinks, and following a link when its cell
 * is clicked. The links themselves live on the sheet (Sheet.hyperlinks); HYPERLINK formulas are followed too.
 */
export class HyperlinkCommands {
  constructor(private readonly c: SheetController) {}

  private get active() {
    const { row, col } = this.c.selection.active
    const merge = this.c.sheet.merges.at(row, col)
    return merge ? { row: merge.r0, col: merge.c0 } : { row, col }
  }

  /** The sheet's own link on a cell (Insert › Link), or null. */
  at(row: number, col: number): Hyperlink | null {
    return hyperlinkAt(this.c.sheet.hyperlinks, row, col)
  }

  /** What a click on the cell follows: its link, else a HYPERLINK formula's target; null for neither. */
  target(row: number, col: number): Omit<Hyperlink, "range"> | null {
    const merge = this.c.sheet.merges.at(row, col)
    if (merge) ({ r0: row, c0: col } = merge)
    return this.at(row, col) ?? formulaLink(this.c.wb.sheets, this.c.sheetIndex, row, col, todaySerial())
  }

  /** The link on the active cell (Edit Link opens with it). */
  activeLink(): Hyperlink | null {
    const { row, col } = this.active
    return this.at(row, col)
  }

  /** Whether the selection has any of the sheet's links (Remove Link, Clear Hyperlinks). */
  selectionHas(): boolean {
    return hyperlinksIn(this.c.sheet.hyperlinks, this.c.selection.ranges).length > 0
  }

  /**
   * "Text to display" for the dialog: the active cell's text, or null when its content isn't text (a number, a
   * formula), which Excel shows greyed as "<<Selection in Document>>".
   */
  displayText(): string | null {
    const { row, col } = this.active
    const cells = this.c.sheet.cells
    const input = cells.input(row, col)
    if (input?.startsWith("=")) return null
    const value = cells.value(row, col)
    if (value === null || value === undefined) return ""
    return typeof value === "string" ? value : null
  }

  /** Insert or Edit Link's OK: the link on every selected range, the text in the active cell, Hyperlink style. */
  set(init: LinkInit): boolean {
    if (this.c.edit) this.c.commitEdit()
    const sheet = this.c.sheetIndex
    const ranges = this.c.selection.ranges.map((r) => ({ ...r }))
    const link = { address: init.address, location: init.location, tooltip: init.tooltip || undefined }
    const ops: Op[] = [{ kind: "hyperlinks", sheet, links: withHyperlink(this.c.sheet.hyperlinks, ranges, link) }]
    const { row, col } = this.active
    const current = this.displayText()
    if (init.text !== null && current !== null && init.text !== current) {
      const old = this.c.sheet.cells.get(row, col)
      ops.push({ kind: "setCells", sheet, cells: [{ row, col, data: { ...old, input: init.text, value: init.text } }] })
    }
    ops.push({ kind: "format", sheet, ranges, patch: HYPERLINK_STYLE })
    const sel = this.c.selection
    return this.c.transact(ops, { label: "Insert Hyperlink", before: sel, after: sel }) !== null
  }

  /**
   * AutoCorrect's "Internet and network paths with hyperlinks", on in Excel by default: typing a web address
   * (https://…, www.…) or an e-mail address alone in a cell makes it a link with the Hyperlink style.
   */
  autoLinkOps(row: number, col: number, text: string): Op[] {
    const t = text.trim()
    if (!t || t !== text || /\s/.test(t) || t.startsWith("=")) return []
    if (!/^(https?:\/\/|ftp:\/\/|www\.)\S+\.\S+|^mailto:\S+@\S+$|^[^\s@/\\:]+@[^\s@/\\]+\.[^\s@/\\]+$/i.test(t))
      return []
    const sheet = this.c.sheet
    if (sheet.protection && !sheet.protection.insertHyperlinks) return []
    if (hyperlinkAt(sheet.hyperlinks, row, col)) return []
    const range = { r0: row, c0: col, r1: row, c1: col }
    return [
      {
        kind: "hyperlinks",
        sheet: this.c.sheetIndex,
        links: withHyperlink(sheet.hyperlinks, [range], { address: completeAddress(t) }),
      },
      { kind: "format", sheet: this.c.sheetIndex, ranges: [range], patch: HYPERLINK_STYLE },
    ]
  }

  /** The op that removes the links in `ranges` (Clear All uses it too); null when there are none. */
  removalOp(ranges: readonly Range[] = this.c.selection.ranges, sheetIndex = this.c.sheetIndex): Op | null {
    const links = this.c.wb.sheets[sheetIndex].hyperlinks
    if (!hyperlinksIn(links, ranges).length) return null
    return { kind: "hyperlinks", sheet: sheetIndex, links: withHyperlink(copyHyperlinks(links), ranges, null) }
  }

  /**
   * Clear Hyperlinks (`formats` false) takes the links off and leaves the cells as they look; Remove Hyperlinks and
   * the cell menu's Remove Link take the formatting too, as Excel does.
   */
  remove(formats: boolean) {
    if (this.c.edit) this.c.commitEdit()
    const ranges = this.c.selection.ranges
    const op = this.removalOp(ranges)
    if (!op) return
    const ops: Op[] = [op]
    if (formats) {
      // Only the linked cells lose their formatting.
      const linked = hyperlinksIn(this.c.sheet.hyperlinks, ranges).flatMap((l) =>
        ranges.map((r) => intersect(l.range, r)).filter((x): x is Range => x !== null),
      )
      ops.push({ kind: "format", sheet: this.c.sheetIndex, ranges: linked, patch: CLEAR_ALL })
    }
    const sel = this.c.selection
    this.c.transact(ops, { label: formats ? "Remove Hyperlinks" : "Clear Hyperlinks", before: sel, after: sel })
  }

  /** Where the link on a cell goes, with a file's path resolved against the workbook's folder. Null: no link. */
  resolve(row: number, col: number): LinkTarget | null {
    const link = this.target(row, col)
    if (!link) return null
    const kind = hyperlinkKind(link)
    if (kind === "place") return link.location ? { kind: "place", location: link.location } : null
    const address = link.address!.trim()
    if (kind === "web" || kind === "mail") return { kind: "external", target: address, file: false }
    const path = resolvePath(this.c.shell.documentPath?.() ?? null, fileAddress(address))
    if (WORKBOOK.test(path)) return { kind: "workbook", path, ...(link.location ? { location: link.location } : {}) }
    return { kind: "external", target: path, file: true }
  }

  /** Follows the link on a cell (a click, Open Link). Returns false when the cell has no link. */
  follow(row: number, col: number): boolean {
    const target = this.resolve(row, col)
    if (!target) return false
    if (target.kind === "place") {
      this.goTo(target.location)
      return true
    }
    this.c.shell.followLink?.(target)
    return true
  }

  /** Goes to a place in this workbook ("Sheet2!B4", a name); Excel's message when there's no such place. */
  goTo(location: string): boolean {
    const names = this.c.names
    const place = names.resolve(location)
    if (!place) {
      this.c.shell.alert(BAD_REFERENCE)
      return false
    }
    names.goTo(place)
    return true
  }
}

function intersect(a: Range, b: Range): Range | null {
  const r = { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) }
  return r.r0 > r.r1 || r.c0 > r.c1 ? null : r
}

/** A file: URL as a path ("file:///C:/x.pdf" → "C:/x.pdf", "file:///home/alex/x.pdf" → "/home/alex/x.pdf"). */
function fileAddress(address: string): string {
  const m = /^file:\/\/(localhost)?(\/.*)$/i.exec(address)
  if (!m) return address
  let path = m[2]
  try {
    path = decodeURIComponent(path)
  } catch {
    // Not encoded.
  }
  return /^\/[a-z]:[\\/]/i.test(path) ? path.slice(1) : path
}

/**
 * A link's file path made full: Excel keeps a file near the workbook relative to the workbook's folder
 * ("../Reports/Q1.pdf"). Full paths stay as they are; without a saved workbook a relative path can't be resolved.
 */
export function resolvePath(workbook: string | null, path: string): string {
  if (/^([a-z]:[\\/]|\/|\\\\|content:)/i.test(path) || !workbook || workbook.startsWith("content:")) return path
  const windows = /^[a-z]:[\\/]/i.test(workbook) || workbook.startsWith("\\\\")
  const sep = windows ? "\\" : "/"
  const parts = workbook.split(/[\\/]/)
  parts.pop()
  for (const part of path.split(/[\\/]/)) {
    if (part === "..") {
      if (parts.length > 1) parts.pop()
    } else if (part !== "." && part !== "") parts.push(part)
  }
  return parts.join(sep)
}
