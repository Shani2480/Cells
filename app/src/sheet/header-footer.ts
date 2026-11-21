import type { HeaderFooterText } from "@cells/core"
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

const named = (hf: HeaderFooterText): HeaderFooterText => ({
  left: codesToNames(hf.left),
  center: codesToNames(hf.center),
  right: codesToNames(hf.right),
})

/**
 * Page Layout view's header and footer, edited on the page as in Excel: a click in a page's top or bottom margin
 * opens its three sections there, the Header & Footer tab inserts elements and moves between header and footer, and
 * a click back on the cells keeps the text (Esc drops it). The same header and footer print on every page.
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
    const text = named(this.c.sheet.pageSetup[where])
    this.editing = {
      where,
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
    this.c.print.change((s) => (s[where] = { ...value }), where === "header" ? "Header" : "Footer")
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
    const before = this.c.sheet.pageSetup[e.where]
    if (value.left !== before.left || value.center !== before.center || value.right !== before.right)
      this.c.print.change((s) => (s[e.where] = value), e.where === "header" ? "Header" : "Footer")
    this.c.changed()
  }

  /** Esc: stop editing without keeping the text. */
  cancel() {
    if (!this.editing) return
    this.editing = null
    this.c.changed()
  }
}
