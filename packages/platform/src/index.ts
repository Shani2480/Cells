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
  /**
   * Calls back with the id of a menu bar item the person chose (macOS; `app/src/shell/menu-bar.ts`). Absent where
   * there is no menu bar. Returns an unsubscribe function.
   */
  onMenuCommand?(callback: (id: string) => void): () => void
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
  readonly ai: AiPort
}

/** Where an AI request goes: the providers Inkspire uses (LM Studio is an OpenAI-compatible endpoint). */
export type AiTarget = {
  provider: "openai" | "anthropic" | "ollama" | "compatible"
  /** Blank: the provider's default. */
  model?: string
  baseUrl?: string
}

/** An AbortSignal, as far as the AI port uses it (this package has no DOM types). */
export type StopSignal = {
  readonly aborted: boolean
  addEventListener(type: "abort", listener: () => void): void
  removeEventListener(type: "abort", listener: () => void): void
}

export type AiRequest = { system: string; prompt: string; context: string; target: AiTarget }

/** Per provider that takes a key: "saved", "environment" or "missing"; `storage` is "keyring" or "file". */
export type AiKeyStatus = { openai: string; anthropic: string; compatible: string; storage: string }

/**
 * The AI tools' requests (`docs/ai.md`). Keys never reach the webview: the desktop shell keeps them in the system
 * keyring Inkspire uses, and sends the requests itself. A browser has no shell, so there `available` is false.
 */
export interface AiPort {
  readonly available: boolean
  /** Sends a request; `onText` gets the reply as it streams in. Resolves with the whole reply. */
  generate(request: AiRequest, onText: (text: string) => void, signal?: StopSignal): Promise<string>
  /** Checks that the provider answers; resolves with a message to show, rejects with the problem. */
  test(target: AiTarget): Promise<string>
  keyStatus(): Promise<AiKeyStatus | null>
  /** Saves a key (empty removes it); resolves with where it went: "keyring", "file" or "none". */
  saveKey(provider: "openai" | "anthropic" | "compatible", key: string): Promise<string>
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
  /**
   * Opens a link's target with the system: a web address or mailto: in the browser or mail app, a file (full path)
   * in the app the system has for it. Rejects when it can't (a browser tab can't open files).
   */
  external(target: string): Promise<void>
  /**
   * Insert › Object: opens an embedded file in the app the system has for it, from a copy in Cells' cache (a browser
   * downloads it instead). Optional: where it's absent, double-clicking an embedded file says it can't be opened.
   */
  openCopy?(name: string, bytes: Uint8Array): Promise<void>
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
  /**
   * The system's Open dialog for any file, giving only its full path (Insert › Link's Browse for File); absent where
   * a picked file has no lasting path (a browser, Android's content links). Null when cancelled.
   */
  pickPath?(title: string): Promise<string | null>
  /** The system's Save As dialog; null when cancelled. */
  pickSave(suggestedName: string, types: FileType[]): Promise<FileRef | null>
  write(ref: FileRef, bytes: Uint8Array): Promise<void>
  exists(ref: FileRef): Promise<boolean>
  /** Calls back with a file dropped on the window (the first, when several are). Returns an unsubscribe function. */
  onDrop(handler: (file: { ref: FileRef; bytes: Uint8Array }) => void): () => void
}
