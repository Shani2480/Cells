/**
 * What changed in each version, in plain words, newest first. Shown on the What's new page (File → Account, with
 * the backstage in Phase 3) and once after an update. A test checks the newest entry matches the version built.
 */
export type ReleaseNote = { version: string; date: string; notes: string[] }

export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "0.3.0",
    date: "2026-11-14",
    notes: [
      "The ribbon: Home, Insert, Page Layout, Formulas, Data, Review, View and Help, with KeyTips (press Alt), Ctrl+F1 to collapse it, a Quick Access Toolbar you can add commands to, and Search (Alt+Q) to find any command. Commands that aren't built yet are greyed out and say so.",
      "Cell formatting: fonts, colours, fills, borders, alignment, wrap, rotation, indents, merging, cell styles, Format Painter, and the Format Cells dialog (Ctrl+1) with all six tabs.",
      "Number formats work like Excel's, including custom codes, dates and times, fractions and conditions; checked against Excel on 594 cases.",
      "Conditional formatting: highlight rules, top/bottom, data bars, colour scales, icon sets, formula rules, and Manage Rules.",
      "Sheet tabs: add, rename, move and copy (drag, or Ctrl+drag), colour, hide and unhide.",
      "View: freeze panes, zoom (status bar, Ctrl+wheel, pinch), gridlines and headings on and off, and dark mode.",
      "Notes and threaded comments; Protect Sheet and Protect Workbook, with passwords.",
      "The formula bar grows to several lines and can be resized. The Name Box lists and creates names, and fx opens Insert Function.",
      "On a touch screen: tap to select, drag the round handles to stretch the selection, long-press for the menu, double-tap to edit, pinch to zoom. The sheet scrolls so the keyboard doesn't cover the cell you're typing in.",
      "This page opens by itself once after an update.",
    ],
  },
  {
    version: "0.2.0",
    date: "2026-11-14",
    notes: [
      "Formulas calculate: 262 Excel functions, dynamic arrays that spill, circular reference warnings and manual calculation with F9.",
      "References in a formula are coloured, with matching boxes on the grid; point at cells with the mouse or the arrow keys.",
      "While typing a formula, a list suggests functions (Tab inserts one) and a tip shows the function's arguments.",
      "F4 cycles a reference through $A$1, A$1, $A1 and A1. Ctrl+` shows the formulas instead of their results.",
      "Inserting or deleting rows and columns, moving cells and copying formulas update the references, as in Excel.",
      "Calculation runs in the background, so the sheet stays responsive while it works.",
    ],
  },
  {
    version: "0.1.0",
    date: "2026-11-14",
    notes: [
      "The first version: a fast grid of 1,048,576 rows by 16,384 columns with Excel's keyboard shortcuts.",
      "Typing, editing, fill handle, copy and paste, insert and delete, row heights and column widths, hide and unhide.",
      "Undo and redo up to 100 steps.",
    ],
  },
]

export const CURRENT_VERSION = RELEASE_NOTES[0].version

const SEEN_KEY = "cells.seen-version"

/**
 * Whether to open What's New on start: once after an update (a different version was seen before). A first start
 * only records the version. Storage can be missing or blocked; then nothing shows.
 */
export function whatsNewOnStart(storage: Pick<Storage, "getItem" | "setItem"> | null = safeStorage()): boolean {
  if (!storage) return false
  try {
    const seen = storage.getItem(SEEN_KEY)
    storage.setItem(SEEN_KEY, CURRENT_VERSION)
    return seen !== null && seen !== CURRENT_VERSION
  } catch {
    return false
  }
}

function safeStorage(): Storage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage
  } catch {
    return null
  }
}
