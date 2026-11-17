import {
  columnName,
  isChart,
  plainText,
  selectCell,
  selectRange,
  type Range,
  type SheetObject,
  type Workbook,
} from "@cells/core"
import type { SheetController } from "./controller"

/**
 * Review › Check Accessibility: what makes the workbook hard to use with a screen reader or for people with low
 * vision, as Excel's Accessibility Checker groups it (errors, then warnings). Each issue says where it is, and the
 * pane offers Excel's recommended actions for it.
 */

export type IssueKind = "altText" | "tableHeader" | "contrast" | "merged" | "sheetName" | "linkText"
export type Severity = "error" | "warning"

export type Issue = {
  kind: IssueKind
  sheet: number
  /** What the pane lists: "Chart 1", "A1:C1", "Sheet1". */
  label: string
  target: { object: number } | { range: Range } | { sheetOnly: true }
}

/** Excel's names and explanations for each kind of issue (Why Fix, How to Fix). */
export const ISSUE_INFO: Record<IssueKind, { severity: Severity; title: string; why: string; how: string }> = {
  altText: {
    severity: "error",
    title: "Missing alternative text",
    why: "Alternative text helps people who can't see the screen understand what's in charts and other objects.",
    how: "Describe the chart in Alt Text: what it shows and the point it makes. If it's only decoration, mark it as decorative.",
  },
  tableHeader: {
    severity: "error",
    title: "No header row specified",
    why: "Header rows tell people using a screen reader what each column of a table holds.",
    how: "Turn on the table's Header Row (Table Design › Header Row) and give each column a clear name.",
  },
  contrast: {
    severity: "warning",
    title: "Hard-to-read text contrast",
    why: "Text that doesn't stand out from its fill is hard to read for people with low vision or colour blindness.",
    how: "Use a darker or lighter text colour, or another fill, so the text stands out (a contrast of at least 4.5 to 1, or 3 to 1 for large text).",
  },
  merged: {
    severity: "warning",
    title: "Merged cells",
    why: "Screen readers lose their place in merged cells, and keyboard navigation through them is confusing.",
    how: "Unmerge the cells. To keep the look, use Center Across Selection (Format Cells › Alignment) instead.",
  },
  sheetName: {
    severity: "warning",
    title: "Default sheet names",
    why: "A sheet's name tells people moving between sheets with a screen reader what each one holds.",
    how: "Rename the sheet to say what's on it (double-click its tab).",
  },
  linkText: {
    severity: "warning",
    title: "Unclear hyperlink text",
    why: 'Screen reader users often move from link to link; text such as "click here" doesn\'t say where a link goes.',
    how: "Change the link's text to say where it goes (Edit Hyperlink › Text to display).",
  },
}

/** The order the pane lists kinds in. */
export const ISSUE_ORDER: IssueKind[] = ["altText", "tableHeader", "contrast", "merged", "sheetName", "linkText"]

/** Text that says nothing about where a link goes. */
const UNCLEAR_LINK = /^(click here|here|click|link|this link|more|read more|learn more|go|this)\.?$/i

/** At most this many cells per sheet are listed for poor contrast (the rest are summed into the count). */
const CONTRAST_LIST_MAX = 200

/** The WCAG relative luminance of "#rrggbb". */
function luminance(hex: string): number {
  const n = parseInt(hex.slice(1, 7), 16)
  const ch = (shift: number) => {
    const v = ((n >> shift) & 255) / 255
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
  }
  return 0.2126 * ch(16) + 0.7152 * ch(8) + 0.0722 * ch(0)
}

/** The contrast ratio of two "#rrggbb" colours (1 to 21). */
export function contrastRatio(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x)
  return (hi + 0.05) / (lo + 0.05)
}

const isHex = (c: string | undefined): c is string => !!c && /^#[0-9a-f]{6}$/i.test(c)

/** Whether a style's text is hard to read on its fill (Automatic text is black, No Fill is white, as printed). */
function poorContrast(wb: Workbook, styleId: number): boolean {
  const s = wb.styles.get(styleId)
  const text = isHex(s.fontColor) ? s.fontColor : "#000000"
  const fill = isHex(s.fill) ? s.fill : "#ffffff"
  const size = s.fontSize ?? 11
  const large = size >= 18 || (s.bold && size >= 14)
  return contrastRatio(text, fill) < (large ? 3 : 4.5)
}

const rangeLabel = (r: Range) => {
  const cell = (row: number, col: number) => columnName(col) + (row + 1)
  return r.r0 === r.r1 && r.c0 === r.c1 ? cell(r.r0, r.c0) : `${cell(r.r0, r.c0)}:${cell(r.r1, r.c1)}`
}

/**
 * Whether an object needs Alt Text: charts, pictures, shapes and groups, unless described or marked decorative. A
 * shape whose own text says what it is (a text box, WordArt, a labelled shape) reads out already, as in Excel.
 */
function needsAltText(o: SheetObject): boolean {
  if (o.altText?.trim() || o.decorative) return false
  if (!isChart(o) && o.kind === "shape" && plainText(o.text).trim()) return false
  if (!isChart(o) && o.hidden) return false
  return true
}

/** Every issue in the workbook, sheet by sheet (hidden sheets too, as Excel checks them). */
export function checkAccessibility(wb: Workbook): Issue[] {
  const out: Issue[] = []
  wb.sheets.forEach((sheet, index) => {
    for (const o of sheet.objects)
      if (needsAltText(o)) out.push({ kind: "altText", sheet: index, label: o.name, target: { object: o.id } })
    for (const table of sheet.tables)
      if (!table.header)
        out.push({ kind: "tableHeader", sheet: index, label: table.name, target: { range: table.range } })
    const used = sheet.cells.usedBounds()
    if (used) {
      const poor = new Map<number, boolean>()
      let listed = 0
      sheet.cells.forEach(
        used,
        (row, col, data) => {
          if (listed >= CONTRAST_LIST_MAX || data.value === undefined || data.value === "") return
          const id = sheet.styleAt(row, col)
          let bad = poor.get(id)
          if (bad === undefined) poor.set(id, (bad = poorContrast(wb, id)))
          if (!bad) return
          listed++
          const r = { r0: row, r1: row, c0: col, c1: col }
          out.push({ kind: "contrast", sheet: index, label: rangeLabel(r), target: { range: r } })
        },
        true,
      )
    }
    for (const m of sheet.merges.all)
      out.push({ kind: "merged", sheet: index, label: rangeLabel(m), target: { range: { ...m } } })
    const hasContent = used !== null || sheet.objects.length > 0
    if (hasContent && /^Sheet\d+$/i.test(sheet.name))
      out.push({ kind: "sheetName", sheet: index, label: sheet.name, target: { sheetOnly: true } })
    for (const link of sheet.hyperlinks) {
      const text = String(sheet.cells.value(link.range.r0, link.range.c0) ?? link.display ?? "").trim()
      if (UNCLEAR_LINK.test(text))
        out.push({ kind: "linkText", sheet: index, label: rangeLabel(link.range), target: { range: link.range } })
    }
  })
  return out
}

/** Goes to an issue: its sheet, then the chart or cells. */
export function goToIssue(c: SheetController, issue: Issue) {
  if (c.sheetIndex !== issue.sheet) c.activate(issue.sheet)
  const t = issue.target
  if ("object" in t) c.objects.select(t.object)
  else if ("range" in t) {
    c.objects.select(null)
    const r = t.range
    c.setSelection(r.r0 === r.r1 && r.c0 === r.c1 ? selectCell({ row: r.r0, col: r.c0 }) : selectRange(r), {
      row: r.r0,
      col: r.c0,
    })
  }
}

/** A recommended action the pane offers for an issue (Excel's dropdown). */
export type IssueAction = { label: string; run: () => void }

export function issueActions(c: SheetController, issue: Issue, open: (dialog: string) => void): IssueAction[] {
  const go = () => goToIssue(c, issue)
  switch (issue.kind) {
    case "altText": {
      const id = (issue.target as { object: number }).object
      return [
        { label: "Add a description", run: () => (go(), open(`altText:${id}`)) },
        { label: "Mark as decorative", run: () => (go(), c.objects.setAltText(id, "", true)) },
      ]
    }
    case "tableHeader":
      return [{ label: "Use the header row", run: () => (go(), c.tables.setOption("header", true)) }]
    case "contrast":
      return [
        {
          label: "Use black or white text",
          run: () => {
            go()
            const style = c.wb.styles.get(c.sheet.styleAt(c.selection.active.row, c.selection.active.col))
            const fill = isHex(style.fill) ? style.fill : "#ffffff"
            const color = contrastRatio("#000000", fill) >= contrastRatio("#ffffff", fill) ? "#000000" : "#ffffff"
            c.format.apply({ fontColor: color })
          },
        },
      ]
    case "merged":
      return [{ label: "Unmerge", run: () => (go(), c.format.unmerge()) }]
    case "sheetName":
      return [{ label: "Rename sheet", run: () => c.sheets.beginRename(issue.sheet) }]
    case "linkText":
      return [{ label: "Edit link", run: () => (go(), open("insertLink")) }]
  }
}
