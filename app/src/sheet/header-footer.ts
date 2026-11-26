import {
  headerFooterKey,
  headerFooterKind,
  type HeaderFooterKind,
  type HeaderFooterText,
  type PageSetup,
} from "@cells/core"
import type { SheetController } from "./controller"

/** Excel shows codes by name while editing (&[Page]) and keeps them as letters in the file (&P). */
const CODE_NAMES: [string, string][] = [
  ["&P", "&[Page]"],
  ["&N", "&[Pages]"],
  ["&D", "&[Date]"],
  ["&T", "&[Time]"],
  ["&Z", "&[Path]"],
  ["&F", "&[File]"],
  ["&A", "&[Tab]"],
]
export const codesToNames = (s: string) =>
  s.replace(/&&|&[PNDTZFA]/g, (code) => CODE_NAMES.find(([c]) => c === code)?.[1] ?? code)
export const namesToCodes = (s: string) => CODE_NAMES.reduce((out, [code, name]) => out.split(name).join(code), s)

export type HFWhere = "header" | "footer"
export type HFSection = keyof HeaderFooterText

/** Header & Footer Elements, as the contextual tab's buttons insert them (shown by name). */
export const HF_ELEMENTS = {
  pageNumber: "&[Page]",
  pages: "&[Pages]",
  date: "&[Date]",
  time: "&[Time]",
  path: "&[Path]&[File]",
  file: "&[File]",
  sheet: "&[Tab]",
} as const

export type HFEditing = {
  where: HFWhere
  /** Which of the sheet's headers the page shows: the first page's, an even page's or the usual one. */
  kind: HeaderFooterKind
  /** The page it's edited on (the page's row and column band in Page Layout). */
  rowBand: number
  colBand: number
  /** The three sections as typed, codes shown by name. */
  text: HeaderFooterText
  section: HFSection
  /** Where the caret is in each section (the element buttons insert there). */
  caret: Record<HFSection, number>
  /** Counts the times the caret is put back in the section from outside (the overlay focuses it again). */
  focus: number
}

export type HFOption = keyof Pick<
  PageSetup,
  "differentFirst" | "differentOddEven" | "scaleWithDoc" | "alignWithMargins"
>

const HF_OPTION_LABELS: Record<HFOption, string> = {
  differentFirst: "Different First Page",
  differentOddEven: "Different Odd & Even Pages",
  scaleWithDoc: "Scale with Document",
  alignWithMargins: "Align with Page Margins",
}

/** The name Excel shows by a header being edited when the sheet has more than one ("First Page Header"); null for one. */
export function headerFooterLabel(setup: PageSetup, where: HFWhere, kind: HeaderFooterKind): string | null {
  const noun = where === "header" ? "Header" : "Footer"
  if (kind === "first") return `First Page ${noun}`
  if (kind === "even") return `Even Page ${noun}`
  return setup.differentOddEven ? `Odd Page ${noun}` : setup.differentFirst ? noun : null
}

/**
 * A page's place in the sheet's printout (0 the first), from its row and column band in Page Layout view: the
 * printed page that starts there, or its place in the page order.
 */
export function pagePlace(c: SheetController, rowBand: number, colBand: number): number {
  const p = c.print.current()
  const map = c.pageMap()
  const rb = map?.rows.bands[rowBand]
  const cb = map?.cols.bands[colBand]
  const printed = rb && cb ? p.pages.findIndex((page) => page.rows[0] === rb.first && page.cols[0] === cb.first) : -1
  if (printed >= 0) return printed
  const rows = map?.rows.bands.length ?? 1
  const cols = map?.cols.bands.length ?? 1
  return c.sheet.pageSetup.order === "overThenDown" ? rowBand * cols + colBand : colBand * rows + rowBand
}

const named = (hf: HeaderFooterText): HeaderFooterText => ({
  left: codesToNames(hf.left),
  center: codesToNames(hf.center),
  right: codesToNames(hf.right),
})

/**
 * Page Layout view's header and footer, edited on the page as in Excel: a click in a page's top or bottom margin
 * opens its three sections there, the Header & Footer tab inserts elements and moves between header and footer, and
 * a click back on the cells keeps the text (Esc drops it). With Different First Page or Different Odd & Even Pages
 * on, each page edits the header it prints.
 */
export class HeaderFooterEditor {
  editing: HFEditing | null = null
  private focusCount = 0

  constructor(private readonly c: SheetController) {}

  /**
   * Start editing a page's header or footer, in the section nearest the click. Without a page (Insert › Header &
   * Footer) it's the active cell's page, scrolled into view, in Page Layout view.
   */
  open(where: HFWhere, page?: { rowBand: number; colBand: number }, section: HFSection = "center") {
    if (this.c.edit) this.c.commitEdit()
    if (this.editing) this.commit()
    if (this.c.view !== "pageLayout") this.c.setView("pageLayout")
    if (!page) page = this.activePage(where)
    const kind = headerFooterKind(this.c.sheet.pageSetup, pagePlace(this.c, page.rowBand, page.colBand))
    const text = named(this.c.sheet.pageSetup[headerFooterKey(where, kind)])
    this.editing = {
      where,
      kind,
      ...page,
      text,
      section,
      caret: { left: text.left.length, center: text.center.length, right: text.right.length },
      focus: this.focusCount++,
    }
    this.c.shell.showTab?.("TabHeaderFooter")
    this.c.changed()
  }

  /** The page with the active cell (scrolled so its header or footer shows). */
  private activePage(where: HFWhere): { rowBand: number; colBand: number } {
    const map = this.c.pageMap()
    if (!map) return { rowBand: 0, colBand: 0 }
    const { row, col } = this.c.selection.active
    const find = (bands: { first: number; last: number }[], at: number) =>
      Math.max(
        0,
        bands.findIndex((b) => at >= b.first && at <= b.last),
      )
    const rowBand = find(map.rows.bands, row)
    const colBand = find(map.cols.bands, col)
    const band = map.rows.bands[rowBand]
    if (band) {
      // The header at the page's top, the footer at its foot.
      const top = band.start
      const bottom = band.start + map.rows.page
      const view = this.c.paneHeight
      if (where === "header" && (top < this.c.scrollY || top > this.c.scrollY + view - 60))
        this.c.scrollY = Math.max(0, top - 16)
      if (where === "footer" && (bottom > this.c.scrollY + view || bottom < this.c.scrollY + 60))
        this.c.scrollY = Math.max(0, bottom - view + 16)
    }
    return { rowBand, colBand }
  }

  /** What a section now holds (typing in it). */
  type(section: HFSection, value: string, caret = value.length) {
    const e = this.editing
    if (!e) return
    this.editing = { ...e, section, text: { ...e.text, [section]: value }, caret: { ...e.caret, [section]: caret } }
    this.c.changed()
  }

  /** The caret moved in a section (or it took the focus). */
  place(section: HFSection, caret: number) {
    const e = this.editing
    if (!e || (e.section === section && e.caret[section] === caret)) return
    this.editing = { ...e, section, caret: { ...e.caret, [section]: caret } }
  }

  /** Header & Footer Elements: a code at the caret of the section being edited. */
  insert(name: string) {
    const e = this.editing
    if (!e) return
    const s = e.section
    const at = Math.min(e.caret[s], e.text[s].length)
    const value = e.text[s].slice(0, at) + name + e.text[s].slice(at)
    this.type(s, value, at + name.length)
    if (this.editing) this.editing = { ...this.editing, focus: this.focusCount++ }
  }

  /** Header & Footer › Header or Footer: one of Excel's ready-made ones, kept at once (as Excel does). */
  usePreset(where: HFWhere, value: HeaderFooterText) {
    const e = this.editing
    // The other one's typed text is kept first; this one's is replaced by the ready-made one.
    if (e && e.where !== where) this.commit()
    this.editing = null
    const key = headerFooterKey(where, e?.kind ?? "odd")
    this.c.print.change((s) => (s[key] = { ...value }), where === "header" ? "Header" : "Footer")
    if (e) this.open(where, { rowBand: e.rowBand, colBand: e.colBand }, e.section)
  }

  /** Navigation › Go to Header or Go to Footer (the text typed so far is kept). */
  goTo(where: HFWhere) {
    const e = this.editing
    if (!e || e.where === where) return
    this.open(where, { rowBand: e.rowBand, colBand: e.colBand }, e.section)
  }

  /** Keep what was typed (one undo step) and stop editing. */
  commit() {
    const e = this.editing
    if (!e) return
    this.editing = null
    const value = {
      left: namesToCodes(e.text.left),
      center: namesToCodes(e.text.center),
      right: namesToCodes(e.text.right),
    }
    const key = headerFooterKey(e.where, e.kind)
    const before = this.c.sheet.pageSetup[key]
    if (value.left !== before.left || value.center !== before.center || value.right !== before.right)
      this.c.print.change((s) => (s[key] = value), e.where === "header" ? "Header" : "Footer")
    this.c.changed()
  }

  /**
   * Header & Footer › Options: Different First Page, Different Odd & Even Pages, Scale with Document, Align with Page
   * Margins. The text typed so far is kept first, and the page goes on being edited (as its header may now differ).
   */
  toggleOption(option: HFOption) {
    const e = this.editing
    if (e) this.commit()
    this.c.print.change((s) => (s[option] = !s[option]), HF_OPTION_LABELS[option])
    if (e) this.open(e.where, { rowBand: e.rowBand, colBand: e.colBand }, e.section)
  }

  /** Esc: stop editing without keeping the text. */
  cancel() {
    if (!this.editing) return
    this.editing = null
    this.c.changed()
  }
}
