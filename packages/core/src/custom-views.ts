import type { Range } from "./address"
import type { PageSetup } from "./page-setup"
import type { Sheet } from "./sheet"

/**
 * View › Custom Views: named sets of how each sheet looks (selection, scroll, zoom, frozen panes, gridlines and
 * headings), with the hidden rows and columns and the print settings when the view includes them. Excel keeps them
 * out of undo, and so does Cells: they live beside the workbook, not in its operations.
 */

/** How one sheet looks in a custom view. */
export type CustomSheetView = {
  zoom: number
  /** The first row and column in view (of the scrolling part when panes are frozen). */
  topLeft: { row: number; col: number }
  active: { row: number; col: number }
  selection: Range[]
  freeze: { row: number; col: number; top: number; left: number } | null
  showGridlines: boolean
  showHeadings: boolean
  mode: "normal" | "pageBreakPreview" | "pageLayout"
  /** The hidden rows and columns as [first, last] spans (used when the view includes them). */
  hiddenRows: [number, number][]
  hiddenCols: [number, number][]
  /** The page setup (when the view includes print settings). */
  pageSetup: PageSetup | null
}

export type CustomView = {
  name: string
  /** Excel's id for the view, "{XXXXXXXX-…}"; it ties the sheets' parts and the hidden names to it. */
  guid: string
  printSettings: boolean
  hiddenRowsCols: boolean
  /** The sheet shown when the view is shown. */
  activeSheet: Sheet | null
  sheets: Map<Sheet, CustomSheetView>
}

/** A new view id in Excel's form. */
export function newViewGuid(random: () => number = Math.random): string {
  const hex = (n: number) =>
    Array.from({ length: n }, () =>
      Math.floor(random() * 16)
        .toString(16)
        .toUpperCase(),
    ).join("")
  return `{${hex(8)}-${hex(4)}-4${hex(3)}-${"89AB"[Math.floor(random() * 4)]}${hex(3)}-${hex(12)}}`
}

/** The [first, last] spans of an axis's hidden lines. */
export function hiddenSpans(isHidden: (i: number) => boolean, count: number, last: number): [number, number][] {
  const out: [number, number][] = []
  let start = -1
  const end = Math.min(count - 1, last)
  for (let i = 0; i <= end + 1; i++) {
    const h = i <= end && isHidden(i)
    if (h && start < 0) start = i
    else if (!h && start >= 0) {
      out.push([start, i - 1])
      start = -1
    }
  }
  return out
}
