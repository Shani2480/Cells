import { create } from "zustand"
import { DOWNLOADABLE_DICTIONARIES } from "./dictionaries-data"

/**
 * Spelling languages: English (United States) comes with Cells; the others are Hunspell dictionaries downloaded when
 * the person asks (File › Options › Proofing, the setup, the Spelling dialog), each pinned to a version on jsDelivr
 * and checked against its SHA-256 before it's kept. Kept dictionaries live in the app's own storage (IndexedDB), so
 * they work offline afterwards.
 */

export type DictionaryFile = { size: number; sha256: string }

export type DictionaryEntry = {
  /** A BCP 47 tag, as Excel's dictionary languages: "nl-NL". */
  id: string
  label: string
  /** The word list's licence (SPDX), shown next to the download. */
  license: string
  /** Where index.aff and index.dic are. */
  base: string
  aff: DictionaryFile
  dic: DictionaryFile
}

/** The language that comes with Cells. */
export const BUILT_IN_LANGUAGE = { id: "en-US", label: "English (United States)" } as const

export { DOWNLOADABLE_DICTIONARIES }

export const dictionaryEntry = (id: string) => DOWNLOADABLE_DICTIONARIES.find((d) => d.id === id)

/** A language's name, built in or downloadable ("" for one Cells doesn't know). */
export function languageLabel(id: string): string {
  if (id === BUILT_IN_LANGUAGE.id) return BUILT_IN_LANGUAGE.label
  return dictionaryEntry(id)?.label ?? ""
}

/** A download's size for people ("2.5 MB"). */
export const downloadSize = (d: DictionaryEntry) => `${((d.aff.size + d.dic.size) / 1e6).toFixed(1)} MB`

/**
 * The dictionary that fits a locale: the exact one ("en-GB"), else the first of its language ("nl-BE" → "nl-NL").
 * English locales without their own dictionary get the built-in one; null when there's none for the language.
 */
export function dictionaryForLocale(locale: string): string | null {
  if (!locale) return null
  const exact = [BUILT_IN_LANGUAGE.id, ...DOWNLOADABLE_DICTIONARIES.map((d) => d.id)].find(
    (id) => id.toLowerCase() === locale.toLowerCase(),
  )
  if (exact) return exact
  const lang = locale.split("-")[0].toLowerCase()
  if (lang === "en") return BUILT_IN_LANGUAGE.id
  return DOWNLOADABLE_DICTIONARIES.find((d) => d.id.split("-")[0].toLowerCase() === lang)?.id ?? null
}

// ── Storage ──

type Stored = { id: string; aff: string; dic: string }

const DB = "cells-dictionaries"
const STORE = "dictionaries"

function openDb(): Promise<IDBDatabase | null> {
  if (typeof indexedDB === "undefined") return Promise.resolve(null)
  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB, 1)
      request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: "id" })
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => resolve(null)
    } catch {
      resolve(null)
    }
  })
}

async function withStore<T>(mode: IDBTransactionMode, run: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | null> {
  const db = await openDb()
  if (!db) return null
  return new Promise((resolve) => {
    try {
      const request = run(db.transaction(STORE, mode).objectStore(STORE))
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => resolve(null)
    } catch {
      resolve(null)
    } finally {
      db.close()
    }
  })
}

/** A downloaded dictionary's files (null: not downloaded). */
export async function storedDictionary(id: string): Promise<{ aff: string; dic: string } | null> {
  const got = await withStore<Stored | undefined>("readonly", (s) => s.get(id))
  return got ? { aff: got.aff, dic: got.dic } : null
}

/** The downloaded languages, and those downloading now, as the UI shows them. */
type DictionaryState = {
  installed: string[]
  /** Downloading: language → share done (0–1). */
  downloading: Record<string, number>
  /** The last download that failed, and why. */
  failed: { id: string; message: string } | null
}

export const useDictionaries = create<DictionaryState>(() => ({ installed: [], downloading: {}, failed: null }))

/** Reads which dictionaries are kept (once at start-up). */
export async function loadInstalled(): Promise<void> {
  const keys = await withStore<IDBValidKey[]>("readonly", (s) => s.getAllKeys())
  useDictionaries.setState({ installed: (keys ?? []).map(String).filter((id) => dictionaryEntry(id)) })
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as Uint8Array<ArrayBuffer>)
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("")
}

export const DICTIONARY_MESSAGES = {
  offline: "The dictionary couldn't be downloaded. Check that you're online, then try again.",
  changed: "The downloaded dictionary isn't the one Cells expects, so it wasn't kept. Try again later.",
  storage: "There's no room to keep the dictionary on this device.",
}

/**
 * Downloads a language's dictionary, checks both files against their pinned hashes and keeps them. `fetch` is the
 * platform's (the desktop and Android apps fetch it themselves, a browser directly).
 */
export async function downloadDictionary(
  id: string,
  fetch: (url: string) => Promise<{ bytes: Uint8Array }>,
): Promise<boolean> {
  const entry = dictionaryEntry(id)
  const state = useDictionaries.getState()
  if (!entry || id in state.downloading) return false
  const progress = (share: number) =>
    useDictionaries.setState((s) => ({ downloading: { ...s.downloading, [id]: share } }))
  const done = (failed: DictionaryState["failed"]) =>
    useDictionaries.setState((s) => {
      const downloading = { ...s.downloading }
      delete downloading[id]
      return { downloading, failed }
    })
  progress(0)
  const files: Record<"aff" | "dic", string> = { aff: "", dic: "" }
  const total = entry.aff.size + entry.dic.size
  let got = 0
  for (const part of ["aff", "dic"] as const) {
    let bytes: Uint8Array
    try {
      bytes = (await fetch(`${entry.base}index.${part}`)).bytes
    } catch {
      done({ id, message: DICTIONARY_MESSAGES.offline })
      return false
    }
    if (bytes.length !== entry[part].size || (await sha256(bytes)) !== entry[part].sha256) {
      done({ id, message: DICTIONARY_MESSAGES.changed })
      return false
    }
    files[part] = new TextDecoder().decode(bytes)
    got += bytes.length
    progress(got / total)
  }
  const kept = await withStore("readwrite", (s) => s.put({ id, ...files } satisfies Stored))
  if (kept === null) {
    done({ id, message: DICTIONARY_MESSAGES.storage })
    return false
  }
  done(null)
  useDictionaries.setState((s) => ({ installed: s.installed.includes(id) ? s.installed : [...s.installed, id] }))
  return true
}

/** A downloaded dictionary removed from the device. */
export async function removeDictionary(id: string): Promise<void> {
  await withStore("readwrite", (s) => s.delete(id))
  useDictionaries.setState((s) => ({ installed: s.installed.filter((x) => x !== id) }))
}

/** The languages Spelling can use now: the built-in one, then those downloaded, by name. */
export function usableLanguages(installed: readonly string[]): { id: string; label: string }[] {
  const more = DOWNLOADABLE_DICTIONARIES.filter((d) => installed.includes(d.id)).map((d) => ({
    id: d.id,
    label: d.label,
  }))
  return [{ ...BUILT_IN_LANGUAGE }, ...more.sort((a, b) => a.label.localeCompare(b.label))]
}
