import type { Theme } from "./theme"
import type { ExternalBook } from "./links"
import type { WorkbookProtection } from "./protection"
import { MAX_COLS, parseCell, parseRange } from "./address"
import type { Node } from "./formula/ast"
import { parseFormula } from "./formula/parser"
import type { Sheet } from "./sheet"

/**
 * Defined names (Formulas › Define Name, the Name Box): a workbook-wide name for a reference or formula. Sheet-scoped
 * names arrive with the Name Manager in Phase 4.
 */
export type DefinedName = {
  name: string
  /** What it refers to, without "=" (e.g. `Sheet1!$A$1:$B$5`). */
  ref: string
  comment?: string
}

/** The workbook's names, reachable from the sheet list so operations can rewrite them like formulas. */
export type SheetList = Sheet[] & {
  names?: NameTable
  /** Protect Workbook's state, held in a box so operations can change it. */
  structure?: { protection: WorkbookProtection | null }
  /** Linked workbooks (Data › Workbook Links). */
  links?: { list: ExternalBook[] }
  /** Page Layout › Themes. */
  look?: { theme: Theme }
}

export class NameTable {
  private list: DefinedName[] = []
  private parsed = new Map<string, Node | null>()

  all(): readonly DefinedName[] {
    return this.list
  }

  get(name: string): DefinedName | undefined {
    const upper = name.toUpperCase()
    return this.list.find((n) => n.name.toUpperCase() === upper)
  }

  /** The name's formula, parsed (cached by its text); null if it doesn't parse. */
  node(name: string): Node | null {
    const def = this.get(name)
    if (!def) return null
    if (!this.parsed.has(def.ref)) {
      const result = parseFormula(def.ref)
      this.parsed.set(def.ref, result.ok ? result.ast : null)
    }
    return this.parsed.get(def.ref) ?? null
  }

  replace(names: readonly DefinedName[]) {
    this.list = names.map((n) => ({ ...n })).sort((a, b) => a.name.localeCompare(b.name))
  }

  copy(): DefinedName[] {
    return this.list.map((n) => ({ ...n }))
  }
}

/**
 * Why a name can't be used, in Excel's words; null when it can. Excel's rules: starts with a letter, underscore or
 * backslash; letters, digits, periods and underscores after; not a cell reference (A1, R1C1); at most 255 characters.
 */
export function nameProblem(name: string, others: readonly DefinedName[] = [], self?: string): string | null {
  const invalid = "The name that you entered is not valid."
  if (!name || name.length > 255) return invalid
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) return invalid
  if (/^[RrCc]$/.test(name)) return invalid
  const cell = parseCell(name)
  if ((cell && cell.col < MAX_COLS) || (!cell && parseRange(name)) || /^[Rr]\d*[Cc]\d*$/.test(name)) return invalid
  const upper = name.toUpperCase()
  if (others.some((n) => n.name.toUpperCase() === upper && n.name.toUpperCase() !== self?.toUpperCase()))
    return "The name already exists. Enter a unique name."
  return null
}

/** Rewrites every name's reference with `rewrite` (null: unchanged); returns the names before, or null if none changed. */
export function rewriteNames(sheets: SheetList, rewrite: (ref: string) => string | null): DefinedName[] | null {
  const table = sheets.names
  if (!table || !table.all().length) return null
  const before = table.copy()
  let changed = false
  const next = before.map((n) => {
    const ref = rewrite(n.ref)
    if (ref === null) return n
    changed = true
    return { ...n, ref }
  })
  if (!changed) return null
  table.replace(next)
  return before
}
