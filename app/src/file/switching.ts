import type { WindowPort } from "@cells/platform"
import type { FileCommands } from "./commands"

/** One row of View › Switch Windows (and a stop for Ctrl+F6). */
export type SwitchEntry = { key: string; title: string; current: boolean; go: () => void }

/**
 * Every window Switch Windows lists, in a fixed order: this window's own (its workbooks and their New Window
 * windows), and on the desktop each other Cells window in its place among them, as Excel lists every workbook's
 * windows.
 */
export async function switchEntries(files: FileCommands, win: WindowPort): Promise<SwitchEntry[]> {
  const here: SwitchEntry[] = files.windowList().map(({ doc, c, title }) => ({
    key: `${doc.id}:${doc.windows.key(c)}`,
    title,
    current: doc === files.doc && c === doc.windows.active,
    go: () => files.showWindow(doc, c),
  }))
  if (!win.listWindows || !win.focusWindow) return here
  let all: Awaited<ReturnType<NonNullable<WindowPort["listWindows"]>>>
  try {
    all = await win.listWindows()
  } catch {
    return here
  }
  const focus = win.focusWindow.bind(win)
  return all.flatMap((w) =>
    w.current
      ? here
      : [{ key: w.id, title: w.title.replace(/ - Cells$/, "") || "Cells", current: false, go: () => void focus(w.id) }],
  )
}

/** Ctrl+F6 / Ctrl+Shift+F6: the next or previous of those; false when there's only one. */
export async function cycleEntries(files: FileCommands, win: WindowPort, step: 1 | -1): Promise<boolean> {
  const list = await switchEntries(files, win)
  if (list.length < 2) return false
  const at = Math.max(
    0,
    list.findIndex((e) => e.current),
  )
  list[(at + step + list.length) % list.length]!.go()
  return true
}
