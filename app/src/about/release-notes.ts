/**
 * What changed in each version, in plain words, newest first. Shown on the What's new page (File → Account, with
 * the backstage in Phase 3) and once after an update. A test checks the newest entry matches the version built.
 */
export type ReleaseNote = { version: string; date: string; notes: string[] }

export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "0.4.0",
    date: "2026-11-14",
    notes: [
      "Open and save Excel workbooks (.xlsx, .xlsm and templates). Files made in Excel come back out with their charts, pictures, macros and other parts Cells doesn't show yet kept as they were.",
      "File › Home, New, Open, Save As and Info, with a Recent list you can pin files to. Ctrl+N, Ctrl+O, Ctrl+S, F12 and Ctrl+W work as in Excel.",
      "Cells asks to save your changes before closing or opening another file, and the title bar shows the workbook's name.",
      "AutoRecover keeps a copy of unsaved work every two minutes. If Cells closes unexpectedly, Document Recovery offers it the next time.",
      "CSV and text files: a .csv opens at once (a semicolon CSV too) and saves back in its own delimiter and encoding. A .txt opens through the Text Import Wizard: delimited or fixed width, the first row, the file's encoding, and each column as General, Text, a date in any order, or skipped. Data › From Text/CSV puts a file on a new sheet.",
      "Save As CSV UTF-8 or Text (Tab delimited), written as Excel writes them; with several sheets Cells asks first and saves the active one.",
      "Drop a workbook or CSV on the window to open it.",
      "Big files show their progress while opening. If a file can't be opened, Cells says why (damaged, password-protected, the old .xls format) and lists anything it left out.",
      "Sort: Sort A to Z and Z to A on the data around the cell (headers stay on top), and the Sort dialog for several levels, cell or font colour, custom lists like Jan, Feb, Mar, case-sensitive sorting and left-to-right. The order is checked against Excel.",
      "Filter (Ctrl+Shift+L): a dropdown on each header with sorting, a search box, a list of values to tick (dates by year, month and day), Text, Number and Date Filters, Top 10, Above Average and Filter by Color. The status bar counts the rows shown; Clear and Reapply are on the Data tab. Filters open and save with Excel files, and give the same rows as Excel.",
      "Find and Replace (Ctrl+F, Ctrl+H, Shift+F4 for the next match): in the sheet or the whole workbook, in formulas, values, notes or comments, with wildcards, Match case and Match entire cell contents, and Find All's list of matches. The dialog stays open while you work in the sheet.",
      "Go To Special (Ctrl+G › Special…): blanks, constants, formulas, notes, current region and array, row and column differences, precedents, dependents, the last cell, visible cells only and conditional formats.",
      'Remove Duplicates (Data tab): pick the columns to compare, with or without a header row; duplicate rows go and the rest move up, compared the way Excel does (12 and 12.0 match, "apple" and "Apple" match).',
      "Text to Columns (Data tab): split a column by a delimiter or at fixed widths, set each new column's format, and choose where the result goes. It splits the way Excel does, formulas included.",
      "The formula bar shows a date as 3/19/2026 and a percentage as 100%, as Excel does, instead of the number underneath.",
      "Data Validation (Data tab): limit what a cell takes (whole numbers, decimals, a list, dates, times, text length or a formula of your own), show a message when the cell is selected, and choose what happens when someone types something else: stop it, warn, or just tell them. A list gets a dropdown beside the cell (Alt+Down opens it). Circle Invalid Data rings the cells that break their rule, and Go To Special finds cells with rules. The rules open and save with Excel files.",
    ],
  },
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
