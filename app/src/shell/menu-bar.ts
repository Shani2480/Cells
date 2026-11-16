import type { BackstagePage } from "./Backstage"

/**
 * The macOS menu bar's own commands (the menu is built in `src-tauri/src/menu.rs`, after Excel for Mac's). Each item
 * sends its id here. Cut, Copy and Paste are the system's own items, which reach the grid as clipboard events, and
 * Hide, Quit, Minimize and Full Screen are the system's too.
 */

export type MenuBarAction =
  /** A ribbon command, by its id: it runs only when the ribbon would let it. */
  | { kind: "command"; id: string }
  /** Done in the text field that has the focus (a dialog's box, the formula bar), as the system does. */
  | { kind: "text"; command: "undo" | "redo" | "selectAll" }
  | { kind: "backstage"; page: BackstagePage }
  | { kind: "dialog"; name: string }
  | { kind: "saveAs" }
  | { kind: "close" }
  | { kind: "selectAll" }
  | { kind: "toggleRibbon" }
  /** Closes the window, which asks about unsaved changes first; Cells ends with its window. */
  | { kind: "quit" }

/** The ids `menu.rs` sends (without its `cmd:` prefix). */
export const MENU_BAR_IDS = [
  "About",
  "Settings",
  "Quit",
  "FileNew",
  "FileOpen",
  "FileClose",
  "FileSave",
  "FileSaveAs",
  "FilePrint",
  "Undo",
  "Redo",
  "PasteSpecialDialog",
  "SelectAll",
  "FindDialogExcel",
  "ReplaceDialog",
  "GoTo",
  "ToggleRibbon",
  "ShowFormulas",
  "FormatCellsDialog",
  "WhatsNew",
] as const

export type MenuBarId = (typeof MENU_BAR_IDS)[number]

/**
 * What a menu item does. `editing` is true while a text field other than the grid has the focus: then Undo, Redo and
 * Select All work on its text (the ribbon's would undo the workbook), as in Excel for Mac.
 */
export function menuBarAction(id: string, editing: boolean): MenuBarAction | null {
  switch (id as MenuBarId) {
    case "About":
    case "WhatsNew":
      return { kind: "backstage", page: "about" }
    case "Settings":
      return { kind: "dialog", name: "options" }
    case "FilePrint":
      return { kind: "backstage", page: "print" }
    case "FileSaveAs":
      return { kind: "saveAs" }
    case "FileClose":
      return { kind: "close" }
    case "ToggleRibbon":
      return { kind: "toggleRibbon" }
    case "Quit":
      return { kind: "quit" }
    case "Undo":
      return editing ? { kind: "text", command: "undo" } : { kind: "command", id }
    case "Redo":
      return editing ? { kind: "text", command: "redo" } : { kind: "command", id }
    case "SelectAll":
      return editing ? { kind: "text", command: "selectAll" } : { kind: "selectAll" }
    case "FileNew":
    case "FileOpen":
    case "FileSave":
    case "PasteSpecialDialog":
    case "FindDialogExcel":
    case "ReplaceDialog":
    case "GoTo":
    case "ShowFormulas":
    case "FormatCellsDialog":
      return { kind: "command", id }
    default:
      return null
  }
}

/**
 * Whether Undo, Redo and Select All belong to a text field: one that keeps its own (a dialog's box, the formula bar),
 * or the grid's input while a cell is being edited. Otherwise they are the workbook's.
 */
export function isTextField(element: Element | null, cellEditing: boolean): boolean {
  if (!element) return false
  if (element.closest("[data-grid-input]")) return cellEditing
  if (element instanceof HTMLTextAreaElement) return true
  if (element instanceof HTMLInputElement)
    return !["checkbox", "radio", "button", "submit", "range", "color", "file"].includes(element.type)
  return element instanceof HTMLElement && element.isContentEditable === true
}
