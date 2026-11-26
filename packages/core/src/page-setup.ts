import type { Range } from "./address"

/**
 * How a sheet prints (Page Layout › Page Setup, the Page Setup dialog, File › Print), as Excel keeps it per sheet:
 * paper, orientation, margins, scaling, what prints and the headers and footers.
 */

/** Margins in inches, as Excel stores them (pageMargins). */
export type PageMargins = { left: number; right: number; top: number; bottom: number; header: number; footer: number }

/** One header or footer: its three sections, in Excel's codes (&P page, &N pages, &D date, &A sheet, &F file…). */
export type HeaderFooterText = { left: string; center: string; right: string }

/**
 * A picture in a header or footer (Header & Footer › Picture, printed where its section has &G), with Format
 * Picture's settings.
 */
export type HeaderPicture = {
  /** The picture, as a data URL. */
  src: string
  /** Its name (the file it came from, without the extension). */
  title: string
  /** The size it prints at, in points, after cropping. */
  width: number
  height: number
  /** Cropped from each side, as a share of the picture's own size (0–1). */
  crop: { top: number; bottom: number; left: number; right: number }
  lockAspect: boolean
  /** Image control › Color. */
  color: "automatic" | "grayscale" | "blackAndWhite"
  /** Image control: 0–1, 0.5 as the picture is (Washout is brightness 0.85 and contrast 0.15). */
  brightness: number
  contrast: number
}

/** Where a header or footer picture belongs: "CH" (centre header), "LF" (left footer), "RHEVEN", "CFFIRST"… */
export type HeaderPictureId = string

export function headerPictureId(
  where: "header" | "footer",
  kind: HeaderFooterKind,
  section: keyof HeaderFooterText,
): HeaderPictureId {
  const side = section === "left" ? "L" : section === "center" ? "C" : "R"
  return side + (where === "header" ? "H" : "F") + (kind === "odd" ? "" : kind.toUpperCase())
}

/** A new picture's settings: as big as it is (in points at 96 dpi), uncropped, its colours as they are. */
export function newHeaderPicture(src: string, title: string, pixelWidth: number, pixelHeight: number): HeaderPicture {
  return {
    src,
    title,
    width: pixelWidth * 0.75,
    height: pixelHeight * 0.75,
    crop: { top: 0, bottom: 0, left: 0, right: 0 },
    lockAspect: true,
    color: "automatic",
    brightness: 0.5,
    contrast: 0.5,
  }
}

export type PageSetup = {
  orientation: "portrait" | "landscape"
  /**
   * Excel's paper size number (1 Letter, 5 Legal, 9 A4…; PAPER_SIZES). Null: the default paper, which Excel takes
   * from the printer (Letter in the US and Canada, A4 nearly everywhere else).
   */
  paperSize: number | null
  margins: PageMargins
  /** Adjust to: percent of normal size (10–400). Used when `fitTo` is null. */
  scale: number
  /** Fit to: pages wide by pages tall (0: automatic, as many as it takes). */
  fitTo: { width: number; height: number } | null
  centerHorizontally: boolean
  centerVertically: boolean
  gridlines: boolean
  headings: boolean
  blackAndWhite: boolean
  /** Page order: down the rows first, then across (Excel's default), or across first. */
  order: "downThenOver" | "overThenDown"
  /** The first page's number (null: automatic, 1 unless the sheet follows others). */
  firstPageNumber: number | null
  /** The header and footer of odd pages (of every page, when the options below are off). */
  header: HeaderFooterText
  footer: HeaderFooterText
  /** Header & Footer › Options › Different First Page: the first page prints `firstHeader` and `firstFooter`. */
  differentFirst: boolean
  /** Different Odd & Even Pages: even pages (by their place in the printout, as Excel counts) print the even ones. */
  differentOddEven: boolean
  firstHeader: HeaderFooterText
  firstFooter: HeaderFooterText
  evenHeader: HeaderFooterText
  evenFooter: HeaderFooterText
  /** Scale with Document: the header and footer's text shrinks and grows with the sheet's print scale. */
  scaleWithDoc: boolean
  /** Align with Page Margins: the left and right sections line up with the page's margins. */
  alignWithMargins: boolean
  /** Header and footer pictures, by where they belong (each prints where its section has &G). */
  pictures: Record<HeaderPictureId, HeaderPicture>
  /** The print area: ranges each printing on pages of their own (null: everything with data). */
  printArea: Range[] | null
  /** Rows (and columns) to repeat on every page: first and last index. */
  titleRows: [number, number] | null
  titleCols: [number, number] | null
  /** Manual page breaks: the first row (column) of each new page. */
  rowBreaks: number[]
  colBreaks: number[]
}

/** Paper sizes Excel offers (en-US list), in inches. */
export const PAPER_SIZES: Record<number, { name: string; width: number; height: number }> = {
  1: { name: "Letter", width: 8.5, height: 11 },
  3: { name: "Tabloid", width: 11, height: 17 },
  5: { name: "Legal", width: 8.5, height: 14 },
  6: { name: "Statement", width: 5.5, height: 8.5 },
  7: { name: "Executive", width: 7.25, height: 10.5 },
  8: { name: "A3", width: 297 / 25.4, height: 420 / 25.4 },
  9: { name: "A4", width: 210 / 25.4, height: 297 / 25.4 },
  11: { name: "A5", width: 148 / 25.4, height: 210 / 25.4 },
  13: { name: "B5 (JIS)", width: 182 / 25.4, height: 257 / 25.4 },
  20: { name: "Envelope #10", width: 4.125, height: 9.5 },
  27: { name: "Envelope DL", width: 110 / 25.4, height: 220 / 25.4 },
  28: { name: "Envelope C5", width: 162 / 25.4, height: 229 / 25.4 },
  34: { name: "Envelope B5", width: 176 / 25.4, height: 250 / 25.4 },
  37: { name: "Envelope Monarch", width: 3.875, height: 7.5 },
}

/** Excel's three margin presets (Page Layout › Margins). */
export const MARGIN_PRESETS: Record<"normal" | "wide" | "narrow", PageMargins> = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
}

const EMPTY_HF: HeaderFooterText = { left: "", center: "", right: "" }

/** A new sheet's page setup (Excel's: the default paper, Normal margins, 100%). */
export function defaultPageSetup(): PageSetup {
  return {
    orientation: "portrait",
    paperSize: null,
    margins: { ...MARGIN_PRESETS.normal },
    scale: 100,
    fitTo: null,
    centerHorizontally: false,
    centerVertically: false,
    gridlines: false,
    headings: false,
    blackAndWhite: false,
    order: "downThenOver",
    firstPageNumber: null,
    header: { ...EMPTY_HF },
    footer: { ...EMPTY_HF },
    differentFirst: false,
    differentOddEven: false,
    firstHeader: { ...EMPTY_HF },
    firstFooter: { ...EMPTY_HF },
    evenHeader: { ...EMPTY_HF },
    evenFooter: { ...EMPTY_HF },
    scaleWithDoc: true,
    alignWithMargins: true,
    pictures: {},
    printArea: null,
    titleRows: null,
    titleCols: null,
    rowBreaks: [],
    colBreaks: [],
  }
}

export function copyPageSetup(p: PageSetup): PageSetup {
  return {
    ...p,
    margins: { ...p.margins },
    fitTo: p.fitTo ? { ...p.fitTo } : null,
    header: { ...p.header },
    footer: { ...p.footer },
    firstHeader: { ...p.firstHeader },
    firstFooter: { ...p.firstFooter },
    evenHeader: { ...p.evenHeader },
    evenFooter: { ...p.evenFooter },
    pictures: Object.fromEntries(
      Object.entries(p.pictures).map(([id, pic]) => [id, { ...pic, crop: { ...pic.crop } }]),
    ),
    printArea: p.printArea ? p.printArea.map((r) => ({ ...r })) : null,
    titleRows: p.titleRows ? [...p.titleRows] : null,
    titleCols: p.titleCols ? [...p.titleCols] : null,
    rowBreaks: [...p.rowBreaks],
    colBreaks: [...p.colBreaks],
  }
}

/** Which header and footer a page prints: the first page's, an even page's, or the usual (odd) ones. */
export type HeaderFooterKind = "odd" | "even" | "first"

/** The field holding a header or footer of a kind ("header", "evenHeader", "firstFooter"…). */
export type HeaderFooterKey = "header" | "footer" | "evenHeader" | "evenFooter" | "firstHeader" | "firstFooter"

export function headerFooterKey(where: "header" | "footer", kind: HeaderFooterKind): HeaderFooterKey {
  if (kind === "odd") return where
  return `${kind}${where === "header" ? "Header" : "Footer"}` as HeaderFooterKey
}

/**
 * The kind of header a page prints, by its place in the sheet's printout (0 the first page), not its number: Excel
 * prints the even header on the second page even when it is numbered 3.
 */
export function headerFooterKind(p: PageSetup, index: number): HeaderFooterKind {
  if (p.differentFirst && index === 0) return "first"
  return p.differentOddEven && index % 2 === 1 ? "even" : "odd"
}

/** The header or footer a page prints. */
export function headerFooterOf(p: PageSetup, where: "header" | "footer", index: number): HeaderFooterText {
  return p[headerFooterKey(where, headerFooterKind(p, index))]
}

/**
 * Where a header's left and right sections reach, in inches from the paper's left edge: the page's margins, or
 * (Align with Page Margins off) 0.75 in from the paper's edges, as Excel prints them.
 */
export function headerFooterEdges(p: PageSetup, paperWidth: number): { left: number; right: number } {
  return p.alignWithMargins
    ? { left: p.margins.left, right: paperWidth - p.margins.right }
    : { left: 0.75, right: paperWidth - 0.75 }
}

/** The paper size to use when a sheet names none: A4, but Letter where Letter is the standard. */
export function defaultPaperSize(locale: string): number {
  return /-(US|CA|MX|PH|CL|CO|VE|GT|CR|PA|DO|PR|SV|NI)$/i.test(locale) || /^en$/i.test(locale) ? 1 : 9
}

/** The paper size a sheet prints on (its own, or the default). */
export const paperOf = (p: PageSetup, fallback = 1) => p.paperSize ?? fallback

/** The paper's size in inches, turned for landscape. */
export function paperInches(p: PageSetup, fallback = 1): { width: number; height: number } {
  const paper = PAPER_SIZES[paperOf(p, fallback)] ?? PAPER_SIZES[1]
  return p.orientation === "landscape"
    ? { width: paper.height, height: paper.width }
    : { width: paper.width, height: paper.height }
}

/**
 * Print area, titles and breaks follow inserted and deleted rows or columns, as Excel's Print_Area name does: a
 * range loses the lines deleted from it and goes when all of it is deleted.
 */
export function shiftPageSetup(p: PageSetup, axis: "rows" | "cols", at: number, count: number): PageSetup {
  const next = copyPageSetup(p)
  const move = (i: number) => (count > 0 ? (i >= at ? i + count : i) : i >= at - count ? i + count : i >= at ? -1 : i)
  // A span keeps what is left of it: its ends move, deleted ends close in on what remains.
  const span = (a: number, b: number): [number, number] | null => {
    if (count > 0) return [move(a), b >= at ? b + count : b]
    const end = at - count
    if (a >= at && b < end) return null
    const a2 = a >= end ? a + count : a >= at ? at : a
    const b2 = b >= end ? b + count : b >= at ? at - 1 : b
    return b2 < a2 ? null : [a2, b2]
  }
  const rows = axis === "rows"
  if (next.printArea) {
    const areas = next.printArea
      .map((r) => {
        const s = rows ? span(r.r0, r.r1) : span(r.c0, r.c1)
        if (!s) return null
        return rows ? { ...r, r0: s[0], r1: s[1] } : { ...r, c0: s[0], c1: s[1] }
      })
      .filter((r): r is Range => r !== null)
    next.printArea = areas.length ? areas : null
  }
  if (rows && next.titleRows) next.titleRows = span(...next.titleRows)
  if (!rows && next.titleCols) next.titleCols = span(...next.titleCols)
  const breaks = (list: number[]) => [...new Set(list.map(move).filter((i) => i > 0))].sort((a, b) => a - b)
  if (rows) next.rowBreaks = breaks(next.rowBreaks)
  else next.colBreaks = breaks(next.colBreaks)
  return next
}
