/**
 * What another app (Inkspire) or the system asks Cells to open: a `cells://` link or a workbook's path.
 *
 *   cells://open?file=<path>&range=<range>
 *   cells://edit?file=<path>&for=<document name>
 *
 * `file` is a full path, percent-encoded; `range` is optional, as the Name Box takes it ("Sheet1!A1:D10", a name)
 * or in the R1C1 form Word's LINK fields use ("Sheet1!R1C1:R10C4"). `edit` is Inkspire's Edit Data: the workbook of a
 * chart in the document `for` names, which Cells saves back as it changes. See docs/app-links.md.
 */
export type OpenRequest = { path: string; range?: string; editFor?: string }

/** A request from one launch argument or link; null for anything Cells doesn't open. */
export function parseOpenRequest(item: string): OpenRequest | null {
  const text = item.trim()
  if (!text) return null
  if (!/^cells:/i.test(text)) return isFullPath(text) ? { path: text } : null
  let url: URL
  try {
    url = new URL(text)
  } catch {
    return null
  }
  // "cells://open?…": the action is the host; "cells:open?…" (no slashes) too.
  const action = (url.host || url.pathname.replace(/^\/+/, "")).toLowerCase()
  if (action !== "open" && action !== "edit") return null
  const path = url.searchParams.get("file")?.trim()
  if (!path || !isFullPath(path)) return null
  if (action === "edit") {
    // Only a workbook: a chart's data is never a CSV.
    if (!/\.xlsx$/i.test(path)) return null
    return { path, editFor: url.searchParams.get("for")?.trim() || "Inkspire" }
  }
  const range = url.searchParams.get("range")?.trim()
  return range ? { path, range: a1Range(range) } : { path }
}

/** A link to a file and range, as Cells writes it (Copy Link, and for Inkspire's linked tables). */
export function openLink(path: string, range?: string): string {
  const params = new URLSearchParams({ file: path })
  if (range) params.set("range", range)
  return `cells://open?${params.toString().replace(/\+/g, "%20")}`
}

/** A full path, or an Android document link (a file from "Open with"). */
function isFullPath(path: string): boolean {
  return (
    path.startsWith("content://") || path.startsWith("/") || /^[A-Za-z]:[\\/]/.test(path) || path.startsWith("\\\\")
  )
}

const columnName = (n: number) => {
  let s = ""
  for (let c = n; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}

/** "Sheet1!R1C1:R10C4" → "Sheet1!A1:D10"; anything else as it is. */
export function a1Range(range: string): string {
  const bang = range.lastIndexOf("!")
  const sheet = bang >= 0 ? range.slice(0, bang + 1) : ""
  const ref = range.slice(bang + 1)
  const m = /^R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(ref)
  if (!m) return range
  const cell = (r: string, c: string) => `${columnName(Number(c))}${Number(r)}`
  return sheet + cell(m[1], m[2]) + (m[3] ? `:${cell(m[3], m[4])}` : "")
}
