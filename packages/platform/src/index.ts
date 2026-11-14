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
  /**
   * Asks before the window closes: `allow` resolves false to keep it open (unsaved changes, Cancel). A browser tab
   * can only show its own generic prompt, so there `allow` is called and a false answer shows that prompt.
   */
  onCloseRequested(allow: () => Promise<boolean> | boolean): () => void
  /** Calls back when the window is maximized or restored. Returns an unsubscribe function. */
  onMaximizedChange(callback: (maximized: boolean) => void): () => void
}

export type Platform = {
  /** "linux" | "macos" | "windows" | "android" | "web": for shortcuts and small layout differences only. */
  readonly os: "linux" | "macos" | "windows" | "android" | "web"
  readonly clipboard: ClipboardPort
  readonly window: WindowPort
  readonly files: FilePort
}

/** A file the person picked: enough to read or write it again (a path on the desktop, a handle in a browser). */
export type FileRef = {
  /** The file's name with its extension ("Budget.xlsx"); never a full path. */
  name: string
  /** Desktop and Android: the path or content URI. */
  path?: string
  /** Browsers with the File System Access API: the FileSystemFileHandle. */
  handle?: unknown
}

export type FileType = { name: string; extensions: string[] }

export interface FilePort {
  /** Whether Save writes over the file it opened (false: each save downloads a copy). */
  readonly saveInPlace: boolean
  /** Whether a file in the Recent list can be opened again after a restart. */
  readonly remembersFiles: boolean
  /** The system's Open dialog; null when cancelled. */
  pickOpen(types: FileType[]): Promise<{ ref: FileRef; bytes: Uint8Array } | null>
  read(ref: FileRef): Promise<Uint8Array>
  /** The system's Save As dialog; null when cancelled. */
  pickSave(suggestedName: string, types: FileType[]): Promise<FileRef | null>
  write(ref: FileRef, bytes: Uint8Array): Promise<void>
  exists(ref: FileRef): Promise<boolean>
  /** Calls back with a file dropped on the window (the first, when several are). Returns an unsubscribe function. */
  onDrop(handler: (file: { ref: FileRef; bytes: Uint8Array }) => void): () => void
}
