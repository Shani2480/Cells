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
  readonly settings: SettingsPort
  readonly opens: OpenPort
  readonly links: LinkPort
}

/**
 * Inkspire's linked tables and charts (`docs/linked-ranges.md`), through files in the folder the apps share, so
 * neither app has to start the other: Inkspire lists the ranges its open documents link to, Cells writes what
 * those ranges hold whenever they change.
 */
export interface LinkPort {
  /** The text of `links/wanted.json`; null when there is none. */
  wanted(): Promise<string | null>
  /** Writes `links/<name>.json` (a workbook's linked ranges now); `name` is 16 hex digits. */
  publish(name: string, text: string): Promise<void>
  /** Whether a path is a chart's workbook Inkspire copied out for Edit Data (in its edit folder). */
  isEditCopy(path: string): Promise<boolean>
}

/**
 * Links and files the system hands to Cells: `cells://` links (from Inkspire, a browser) and workbooks opened from
 * the file manager, at start or while Cells runs (one Cells at a time; a second launch passes its request on).
 */
export interface OpenPort {
  /** Calls back with each batch of requests (links, or full paths), the ones from start first. Returns unsubscribe. */
  onRequest(handler: (items: string[]) => void): () => void
}

/**
 * What the desktop session says about the person and their region, for the first-run setup. Every field is blank
 * when the system doesn't give it. The same shape as Inkspire's `system_defaults`.
 */
export type SystemFacts = {
  user?: string
  full_name?: string
  lang?: string
  lc_all?: string
  lc_time?: string
  lc_numeric?: string
  /** `navigator.language`, used only when the environment says nothing. */
  intlLocale?: string
}

/**
 * The settings file Cells shares with Inkspire (`docs/shared-settings.md`): on the desktop
 * `<config dir>/inkspire/settings.json`, in a browser a localStorage entry. Read and written as JSON text; the app
 * merges its own keys into what is there, so the other app's keys survive.
 */
export interface SettingsPort {
  /** The file's text; null when there is none yet (or it can't be read). */
  read(): Promise<string | null>
  write(text: string): Promise<void>
  /** Where the file is, for File › Options ("Shared with Inkspire: ~/.config/inkspire/settings.json"); "" if none. */
  location(): Promise<string>
  systemFacts(): Promise<SystemFacts>
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
