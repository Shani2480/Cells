/**
 * What the shells need from the platform, as interfaces only. `app/src/platform-tauri` implements them for the
 * desktop and Android builds, `app/src/platform-web` for the browser dev build and tests. Nothing here may depend on
 * a particular platform (PLAN.md §2).
 */

/** Rich clipboard content: Excel and other apps read `text` (TSV) and `html`. */
export type ClipboardContent = {
  text: string
  html?: string
}

export interface ClipboardPort {
  /** Writes text (and HTML where the platform can). Used by menu commands; keyboard copy uses the native event. */
  write(content: ClipboardContent): Promise<void>
  /** Reads the clipboard for a menu Paste; null when the platform refuses. */
  read(): Promise<ClipboardContent | null>
}

export interface WindowPort {
  /** False in a browser tab (no title bar buttons are drawn). */
  readonly controllable: boolean
  minimize(): Promise<void>
  toggleMaximize(): Promise<void>
  close(): Promise<void>
  isMaximized(): Promise<boolean>
  setTitle(title: string): Promise<void>
  /** Calls back when the window is maximized or restored. Returns an unsubscribe function. */
  onMaximizedChange(callback: (maximized: boolean) => void): () => void
}

export type Platform = {
  /** "linux" | "macos" | "windows" | "android" | "web": for shortcuts and small layout differences only. */
  readonly os: "linux" | "macos" | "windows" | "android" | "web"
  readonly clipboard: ClipboardPort
  readonly window: WindowPort
}
