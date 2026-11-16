/**
 * Android's documents, through MainActivity's `window.CellsFiles` (src-tauri/gen/android/.../MainActivity.kt). The
 * system's picker hands out `content://` links with read and write access that lasts, so Save writes back to the
 * file and Recent opens it again; files from other apps ("Open with", Share) arrive the same way. Reading and writing
 * a link goes through the fs plugin, which opens it with the content resolver.
 */

type Bridge = {
  displayName(uri: string): string | null
  exists(uri: string): boolean
  openDocument(mimeTypesJson: string): void
  createDocument(name: string, mimeType: string): void
  takeOpened(): string
  setBarColor?(color: string): void
}

type Host = { CellsFiles?: Bridge; __cellsPicked?: (uri: string | null) => void; __cellsOpened?: () => void }

const host = () => globalThis as Host

/** The bridge, on Android only. */
export const androidFiles = (): Bridge | null => host().CellsFiles ?? null

export const isContentUri = (path: string) => path.startsWith("content://")

const MIME: Record<string, string> = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
  xltm: "application/vnd.ms-excel.template.macroEnabled.12",
  csv: "text/csv",
  tsv: "text/tab-separated-values",
  txt: "text/plain",
  pdf: "application/pdf",
}

export const mimeTypeOf = (extension: string) => MIME[extension.toLowerCase()] ?? "application/octet-stream"

/**
 * The MIME types the Open picker offers for these extensions. Some file managers label a CSV `text/comma-separated-
 * values` or a workbook `application/octet-stream` (Downloads, mail attachments), so text and those are offered too.
 */
export function pickerTypes(extensions: string[]): string[] {
  const types = new Set(extensions.map(mimeTypeOf).filter((type) => type !== "application/octet-stream"))
  if (extensions.some((ext) => ext === "csv" || ext === "tsv")) {
    types.add("text/comma-separated-values")
    types.add("text/plain")
  }
  if (extensions.some((ext) => ext.startsWith("xl"))) types.add("application/octet-stream")
  return [...types]
}

const names = new Map<string, string | null>()

/**
 * The file name of a path or a content link. A document link ends in its encoded id, which for local storage is the
 * path (`primary%3ADocuments%2FReport.xlsx`); other providers give only an id, so the content resolver is asked.
 */
export function fileNameOf(path: string): string {
  if (!isContentUri(path)) return path.split(/[/\\]/).pop() || path
  const bridge = androidFiles()
  if (bridge && !names.has(path)) {
    let name: string | null = null
    try {
      name = bridge.displayName(path)
    } catch {
      // The provider refused: the link's own last part follows.
    }
    names.set(path, name || null)
  }
  const shown = names.get(path)
  if (shown) return shown
  const last = path.slice(path.lastIndexOf("/") + 1)
  let decoded = last
  try {
    decoded = decodeURIComponent(last)
  } catch {
    // A malformed escape: the raw part.
  }
  return decoded.split(/[/:]/).pop() || decoded
}

/** One picker at a time: the answer comes back to `__cellsPicked`. */
function pick(show: (bridge: Bridge) => void): Promise<string | null> {
  const bridge = androidFiles()
  if (!bridge) return Promise.resolve(null)
  return new Promise((resolve) => {
    host().__cellsPicked = (uri) => {
      delete host().__cellsPicked
      resolve(uri)
    }
    show(bridge)
  })
}

export const pickOpenDocument = (extensions: string[]) =>
  pick((bridge) => bridge.openDocument(JSON.stringify(pickerTypes(extensions))))

export const pickSaveDocument = (name: string) =>
  pick((bridge) => bridge.createDocument(name, mimeTypeOf(name.split(".").pop() ?? "")))

/** Calls back with the files other apps handed over, now and each time more come. Returns an unsubscribe function. */
export function onOpenedDocuments(handler: (uris: string[]) => void): () => void {
  const bridge = androidFiles()
  if (!bridge) return () => undefined
  const take = () => {
    let uris: unknown
    try {
      uris = JSON.parse(bridge.takeOpened())
    } catch {
      return
    }
    if (Array.isArray(uris) && uris.length) handler(uris.filter((uri): uri is string => typeof uri === "string"))
  }
  host().__cellsOpened = take
  take()
  return () => {
    if (host().__cellsOpened === take) delete host().__cellsOpened
  }
}

/** Gives Android's status and navigation bars the title bar's colour (it changes with the theme). */
export function syncSystemBars() {
  const bridge = androidFiles()
  if (!bridge?.setBarColor || typeof getComputedStyle !== "function") return
  const color = getComputedStyle(document.documentElement).getPropertyValue("--app-title-fill").trim()
  if (/^#[0-9a-f]{6}$/i.test(color)) bridge.setBarColor(color)
}
