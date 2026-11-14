import type { FileRef } from "@cells/platform"

/**
 * File › Open › Recent. On the desktop the list survives restarts (paths in local storage); in a browser the file
 * handles can't be stored, so the list lasts as long as the page.
 */
export type RecentFile = { name: string; path?: string; handle?: unknown; opened: number; pinned?: boolean }

const KEY = "cells.recent"
export const RECENT_LIMIT = 50

let session: RecentFile[] = []

function load(persist: boolean): RecentFile[] {
  if (!persist) return session
  try {
    const raw = JSON.parse(localStorage.getItem(KEY) ?? "[]") as unknown
    if (!Array.isArray(raw)) return []
    return raw.filter(
      (r): r is RecentFile =>
        typeof r?.name === "string" && typeof r?.path === "string" && typeof r?.opened === "number",
    )
  } catch {
    return []
  }
}

function store(list: RecentFile[], persist: boolean) {
  session = list
  if (!persist) return
  try {
    localStorage.setItem(
      KEY,
      JSON.stringify(list.map(({ name, path, opened, pinned }) => ({ name, path, opened, pinned }))),
    )
  } catch {
    // Storage full or blocked: the list just isn't kept.
  }
}

const same = (a: { path?: string; handle?: unknown; name: string }, b: FileRef) =>
  a.path ? a.path === b.path : a.handle !== undefined && a.handle === b.handle

export function recentFiles(persist: boolean): RecentFile[] {
  return [...load(persist)].sort((a, b) => Number(b.pinned ?? 0) - Number(a.pinned ?? 0) || b.opened - a.opened)
}

/** Puts a file at the top of the list (opened or saved now). */
export function touchRecent(ref: FileRef, persist: boolean, now = Date.now()): RecentFile[] {
  if (!ref.path && !ref.handle) return recentFiles(persist)
  const list = load(persist)
  const old = list.find((r) => same(r, ref))
  const next: RecentFile = { name: ref.name, path: ref.path, handle: ref.handle, opened: now, pinned: old?.pinned }
  store([next, ...list.filter((r) => r !== old)].slice(0, RECENT_LIMIT), persist)
  return recentFiles(persist)
}

export function removeRecent(entry: RecentFile, persist: boolean): RecentFile[] {
  store(
    load(persist).filter((r) => !same(r, entry)),
    persist,
  )
  return recentFiles(persist)
}

export function pinRecent(entry: RecentFile, pinned: boolean, persist: boolean): RecentFile[] {
  store(
    load(persist).map((r) => (same(r, entry) ? { ...r, pinned: pinned || undefined } : r)),
    persist,
  )
  return recentFiles(persist)
}

/** The folder a recent file is in, shortened for the list ("Documents › Budgets"). */
export function folderOf(path: string | undefined): string {
  if (!path || path.startsWith("content://")) return ""
  const parts = path.split(/[/\\]/).filter(Boolean)
  parts.pop()
  return parts.slice(-2).join(" › ")
}
