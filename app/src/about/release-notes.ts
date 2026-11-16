/**
 * What changed in each version, in plain words, newest first. Shown on the What's new page (File → Account, with
 * the backstage in Phase 3) and once after an update. A test checks the newest entry matches the version built.
 */
export type ReleaseNote = { version: string; date: string; notes: string[] }

export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "0.7.0",
    date: "2026-11-16",
    notes: [
      "Cells opens Excel 97-2003 workbooks (.xls): values, formulas, formatting, column widths, merged cells, frozen panes, names and notes. It opens a copy, and saving makes an .xlsx; the .xls stays as it was. On Android, Cells now shows up in Open with for .xls files, and the Open picker no longer greys them out.",
      "The cell menu's Filter and Sort work: filter by the cell's value, colour or font colour, clear one column's filter, reapply, Sort A to Z or Z to A, put the cell's colour on top, and Custom Sort. On a phone they open in the long-press panel.",
      "Paste Special is in the cell menu.",
      "The Office Clipboard (the small arrow by Home › Clipboard): the last 24 things you copied in Cells, to paste any of them again, or all at once.",
      "Show Comments (Review tab): every comment on the sheet in a pane beside it, in reading order. Pick one to go to its cell and reply; New starts one on the selected cell.",
      "Pick From Drop-down List (in the cell menu, or Alt+Down on a cell): choose from the words already typed in the column above and below.",
      "Sort A to Z and Z to A work again on the Data tab and in Home's Sort & Filter menu.",
    ],
  },
  {
    version: "0.6.0",
    date: "2026-11-16",
    notes: [
      "Installers for Linux (.deb, AppImage and Flatpak), Windows and macOS, all made from one release.",
      "Cells registers itself for workbooks (.xlsx, .xlsm and templates), and offers itself for .csv and .tsv files: double-click a workbook in your file manager and it opens in Cells, in the Cells that's already running if there is one.",
      "On a Mac, Cells has its own menu bar (Cells, File, Edit, View, Format, Window, Help) with the usual ⌘ shortcuts. macOS keeps ⌘H, ⌘M and ⌘Q, so Replace is ⌃H and Go To is ⌃G, as in Excel for Mac. Quit asks about unsaved changes first.",
      "Cells on Android tablets, with the full ribbon. Open and Save As use Android's own file picker, so workbooks in Downloads, Drive or OneDrive open and save back where they are, and Recent opens them again later. A workbook you open with Cells from a file manager, mail or Drive opens in Cells. With a keyboard attached, the usual Ctrl shortcuts work, and the on-screen keyboard only comes up once you start typing in a cell.",
      "Cells on Android phones, with a layout of its own: a compact top bar (File, Undo, Redo, Find), bigger column and row headings, and a bottom bar with your sheets and the Home, Insert, Formulas, Data and View tools. Each tool opens the ribbon's commands in a panel from the bottom, which you can switch to any tab and swipe down to close. While you type in a cell, a row of formula keys (= ( ) , : $ + - * / & \" and fx) sits above the keyboard. Dialogs and the File view take the whole screen. A tablet keeps the full ribbon, and a split screen switches between the two as it's resized, without losing anything.",
      "On a phone, pressing and holding a cell (or a row or column heading) opens a panel from the bottom: Cut, Copy and Paste as big buttons, then Insert, Delete, Clear Contents, notes and comments, Format Cells and the rest.",
      "On a touch screen the sheet glides on after a quick swipe, and stops when you touch it.",
      "On Android, Back closes what's open (a panel, a dialog, a menu, the File view or the cell you're typing in), and with nothing open it puts Cells in the background with your workbook as it was.",
    ],
  },
  {
    version: "0.5.0",
    date: "2026-11-16",
    notes: [
      "A short setup the first time Cells opens: your name (notes and comments are signed with it), how you write dates and numbers, and light or dark. It starts from what your system says, or from what you already set in Inkspire, and Skip keeps those.",
      "File › Options: your user name and initials, the Office Theme (your system's setting, White or Black), and on the Advanced page your region, short date and separators. Run setup again is there too.",
      "Dates and numbers follow your region: in a day-first region 31/1/2026 is a date, with a decimal comma 1,5 is a number, and cells show dates, decimals and thousands your way. Formulas keep Excel's English form, so files mean the same in Excel anywhere.",
      "Only one Cells runs at a time. A cells:// link (Inkspire's coming \"Edit in Cells\" uses one) or a workbook opened from your file manager opens in the Cells that's already running and comes to the front. A link can point at a range, and Cells selects it.",
      "Copy keeps the formatting: a range pastes into Inkspire, Word or mail as a formatted table (fonts, colours, fills, borders, widths and merged cells), and remembers where it came from so Inkspire can link to it.",
      "Charts can be copied (Ctrl+C, or cut with Ctrl+X). In Inkspire the chart pastes as a chart you can still edit; other apps get a picture. Pasting it back into Cells makes a second chart of the same data.",
      "Ready for Inkspire's next update: tables it links from a workbook follow your changes in Cells as you make them.",
      "Also for Inkspire's next update: Edit Data on an Inkspire chart opens the chart's numbers in Cells (\"Chart in Report.docx\"); every change goes back into the document by itself.",
      "AI tools, if you want them (off until you turn them on in the setup or File › Options › AI): Home › AI Tools writes a formula from your words, explains a formula, suggests a chart or summarizes your data. Use an AI on your own computer (Ollama, LM Studio) or a service you have a key for; Cells shares the choice and the keys with Inkspire, and always says where your cells are sent.",
      "These settings are kept in a file Cells shares with Inkspire, so they're set once for both. Cells picks up a change the next time you switch to it. Inkspire starts using the file with its next update.",
    ],
  },
  {
    version: "0.4.0",
    date: "2026-11-15",
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
      "Tables: Insert › Table (Ctrl+T) or Home › Format as Table turns your data into a table with one of Excel's 60 styles, a filter button on every header and banded rows. The Table Design tab renames it, resizes it, adds a total row (pick Sum, Average, Count and more from each total's dropdown), turns the header row, banded rows and columns and first or last column on and off, and converts it back to a normal range. Formulas can name the table's columns (=SUM(Sales[Amount]), [@Qty]*[@Price]) and follow when you rename them. The table grows as you type beside or under it, Tab on its last cell adds a row, and a formula typed into an empty column fills the whole column. Tables open and save with Excel files.",
      "Flash Fill (Ctrl+E, on the Data tab and under Home › Fill): type one or two examples of what you want next to your data, such as first names from full names, initials, an email address made from two columns or the digits of a phone number, and Cells fills in the rest of the column the same way.",
      "Charts: Insert › Charts (or Alt+F1) draws a chart of the data around the selected cell: column, bar, line, area, pie and doughnut, scatter and more, with Recommended Charts first. Drag a chart to move it, drag its handles to size it (Shift keeps its shape), Delete removes it. The Chart Design tab adds titles, labels, legends and gridlines, applies quick layouts, colours and styles, changes the chart type, switches rows and columns and picks the data with Select Data. Charts redraw as you change the numbers, follow inserted rows and renamed sheets, and open and save with Excel files.",
      "Printing: File › Print (Ctrl+P) shows each page as it will print, with the active sheet, the whole workbook or the selection, a page range, orientation, paper size, margins and scaling. Page Layout sets margins, orientation, paper size, the print area, page breaks, titles that repeat on every page, Fit to a number of pages, and printed gridlines and headings. The Page Setup dialog adds headers and footers (page numbers, date, file and sheet name) and centring on the page. After printing or setting up the page, dashed lines on the sheet show where the pages break. All of it opens and saves with Excel files.",
      "File › Export saves the sheet, the workbook or the selection as a PDF.",
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
