import type { KeyInput } from "./controller"

type KeyEvent = { key: string; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; altKey: boolean }

/**
 * Excel for Mac's shortcuts that use Control where Windows uses Ctrl, because macOS keeps ⌘ with the letter for
 * itself (⌘H hides the app, ⌘G finds the next match): Replace ⌃H, Go To ⌃G, Show Formulas ⌃`.
 */
export const MAC_CONTROL_KEYS = new Set(["h", "g", "`"])

/**
 * ⌘ shortcuts that aren't Cells' on macOS: the menu bar's Hide, Minimize, Quit and Settings, window cycling, and ⌘G
 * (find next elsewhere; Go To is ⌃G).
 */
const MAC_SYSTEM_KEYS = new Set(["h", "m", "q", ",", "`", "g"])

/**
 * The key as Cells' shortcuts read it: the primary modifier is Ctrl, or ⌘ on macOS, where Control also counts for
 * the keys in MAC_CONTROL_KEYS.
 */
export function keyInputFrom(event: KeyEvent, os: string): KeyInput {
  const key = event.key
  const mac = os === "macos"
  const control = mac && event.ctrlKey && !event.metaKey && MAC_CONTROL_KEYS.has(key.toLowerCase())
  return {
    key,
    primary: mac ? event.metaKey || control : event.ctrlKey,
    shift: event.shiftKey,
    alt: event.altKey,
  }
}

/** A ⌘ shortcut that belongs to macOS: the grid and formula bar leave it alone, so the menu bar gets it. */
export function isSystemShortcut(event: KeyEvent, os: string): boolean {
  return os === "macos" && event.metaKey && !event.ctrlKey && MAC_SYSTEM_KEYS.has(event.key.toLowerCase())
}
