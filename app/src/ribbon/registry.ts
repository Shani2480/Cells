import type { FileCommands } from "@/file/commands"
import type { ReactNode } from "react"
import type { SheetController } from "@/sheet/controller"

/**
 * One registry of commands, keyed by Excel's control ids (idMso). The ribbon tabs, the Quick Access Toolbar, Tell Me
 * search and KeyTips all look commands up here, so a command behaves and is named the same everywhere.
 */

/** What commands can reach: the sheet and the app shell around it. */
export type RibbonEnv = {
  c: SheetController
  copy(cut: boolean): void
  paste(): void
  /** Opens one of the shell's dialogs ("goto", "insert", "delete", "formatCells:font", …). */
  dialog(name: string): void
  /** Ctrl+F1: shows or hides the ribbon's commands. */
  toggleRibbon(): void
  /** Opens the File view on a page ("info", "about"). */
  backstage(page?: string): void
  /** File › New, Open, Save… (absent in tests that only drive the ribbon). */
  files?: FileCommands
}

export type MenuEntry =
  /** A command id. */
  | string
  /** A separator. */
  | "-"
  | { heading: string }
  /** Custom content (colour palettes, galleries); `close` closes the menu. */
  | { panel: (env: RibbonEnv, close: () => void) => ReactNode }

export type Command = {
  id: string
  label: string
  /** A name that follows the state (Protect Sheet / Unprotect Sheet, New Note / Edit Note). */
  labelFor?: (env: RibbonEnv) => string
  /** One line for the tooltip. */
  description?: string
  /** As shown in tooltips and menus (Ctrl is shown as ⌘ on macOS). */
  shortcut?: string
  /** The letters Alt shows on the control (Alt, H, then this). */
  keytip?: string
  icon?: ReactNode | ((env: RibbonEnv) => ReactNode)
  /** The picture at 32 px for large buttons, when it differs from `icon`. */
  largeIcon?: ReactNode | ((env: RibbonEnv) => ReactNode)
  run?: (env: RibbonEnv) => void
  enabled?: (env: RibbonEnv) => boolean
  pressed?: (env: RibbonEnv) => boolean
  /** A dropdown; with `run` too, the control is a split button. */
  menu?: MenuEntry[] | ((env: RibbonEnv) => MenuEntry[])
  /** For combo boxes: the text shown and what typing a value does. */
  field?: {
    value: (env: RibbonEnv) => string
    commit: (env: RibbonEnv, text: string) => void
    options?: (env: RibbonEnv) => { label: string; value: string; preview?: ReactNode }[]
    width: number
  }
  /** Shown in place of the command (checkbox controls on View and Page Layout). */
  check?: true
}

/** Commands Excel greys out while sheets are grouped: features that belong to one sheet. */
const ONE_SHEET_PREFIXES = [
  "Table",
  "FormatAsTable",
  "Chart",
  "ConditionalFormatting",
  "CF",
  "DataValidation",
  "Sort",
  "Filter",
  "AdvancedFilter",
  "Review",
  "InsertNewComment",
  "InsertLink",
  "WhatIf",
  "Scenario",
  "GoalSeek",
  "DataTable",
  "Forecast",
  "Consolidate",
  "OutlineSubtotals",
  "RemoveDuplicates",
  "ConvertTextToTable",
  "QuickCode",
  "SheetRename",
  "SheetMoveOrCopy",
  "SheetProtect",
  "SheetBackground",
  "ClearComments",
  "ClearNotes",
  "ClearHyperlinks",
  "RemoveHyperlinks",
  "GetPowerQuery",
  "RefreshAll",
  "AI",
]

/** Ids that match a prefix above but work on a group (Home › Delete's id starts with "Table"). */
const GROUP_OK = new Set(["TableDeleteRowsAndColumnsMenu"])

export function oneSheetOnly(id: string): boolean {
  return !GROUP_OK.has(id) && ONE_SHEET_PREFIXES.some((p) => id.startsWith(p))
}

/** Commands that exist in Excel but not yet here: shown disabled, with this note in the tooltip. */
export const NOT_BUILT = "Not built yet."

export class CommandRegistry {
  private readonly map = new Map<string, Command>()
  /** Commands made on demand (one per defined name in Use in Formula). */
  fallback?: (id: string) => Command | undefined

  add(...commands: Command[]) {
    for (const command of commands) this.map.set(command.id, command)
    return this
  }

  get(id: string): Command | undefined {
    return this.map.get(id) ?? this.fallback?.(id)
  }

  all(): Command[] {
    return [...this.map.values()]
  }

  /** Whether the command does something: it runs, opens a menu or takes a value. */
  isBuilt(id: string): boolean {
    const command = this.get(id)
    return !!command && !!(command.run || command.menu || command.field)
  }

  isEnabled(id: string, env: RibbonEnv): boolean {
    const command = this.get(id)
    if (!command || !this.isBuilt(id)) return false
    if (env.c.group?.active && oneSheetOnly(id)) return false
    return command.enabled ? command.enabled(env) : true
  }
}

export const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform)

/** "Ctrl+B" as the platform shows it. */
export function shortcutText(shortcut: string | undefined): string | undefined {
  if (!shortcut || !isMac) return shortcut
  return (
    shortcut
      // Excel for Mac keeps Control for these (⌘H hides the app): Replace ⌃H, Go To ⌃G, Show Formulas ⌃`.
      .replace(/^Ctrl\+([HhGg`])$/, "⌃$1")
      .replace(/Ctrl\+/g, "⌘")
      .replace(/Shift\+/g, "⇧")
      .replace(/Alt\+/g, "⌥")
  )
}

export function iconOf(command: Command | undefined, env: RibbonEnv, large = false): ReactNode {
  if (!command) return null
  // A large-only picture also serves at 16 px (the stylesheet sizes it) when the group goes compact.
  const icon = large ? (command.largeIcon ?? command.icon) : (command.icon ?? command.largeIcon)
  return typeof icon === "function" ? icon(env) : icon
}
