import { createHash } from "node:crypto"
import { readFileSync } from "node:fs"
import { join } from "node:path"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const root = join(__dirname, "../../..")
const aff = readFileSync(join(root, "node_modules/dictionary-en/index.aff"))
const dic = readFileSync(join(root, "node_modules/dictionary-en/index.dic"))
const file = (b: Uint8Array) => ({ size: b.length, sha256: createHash("sha256").update(b).digest("hex") })

// One downloadable language, whose files are the bundled English ones (so the test needs no network).
vi.mock("./dictionaries-data", () => ({
  DOWNLOADABLE_DICTIONARIES: [
    {
      id: "nl-NL",
      label: "Dutch",
      license: "BSD-3-Clause",
      base: "https://cdn.test/nl/",
      aff: file(aff),
      dic: file(dic),
    },
    { id: "de-DE", label: "German", license: "GPL-2.0", base: "https://cdn.test/de/", aff: file(aff), dic: file(dic) },
  ],
}))

const {
  dictionaryForLocale,
  downloadDictionary,
  loadInstalled,
  removeDictionary,
  storedDictionary,
  usableLanguages,
  useDictionaries,
  DICTIONARY_MESSAGES,
} = await import("./dictionaries")

/** IndexedDB, as much of it as the dictionary store uses, kept in memory. */
function memoryIndexedDb() {
  const rows = new Map<string, unknown>()
  const request = <T>(result: T) => {
    const r: { result: T; onsuccess?: () => void; onerror?: () => void } = { result }
    queueMicrotask(() => r.onsuccess?.())
    return r
  }
  const store = {
    get: (id: string) => request(rows.get(id)),
    put: (v: { id: string }) => request((rows.set(v.id, v), v.id)),
    delete: (id: string) => request((rows.delete(id), undefined)),
    getAllKeys: () => request([...rows.keys()]),
  }
  const db = {
    transaction: () => ({ objectStore: () => store }),
    close: () => undefined,
    createObjectStore: () => store,
  }
  return {
    open: () => {
      const r: { result: typeof db; onsuccess?: () => void; onupgradeneeded?: () => void } = { result: db }
      queueMicrotask(() => r.onsuccess?.())
      return r
    },
  }
}

const files: Record<string, Uint8Array> = { "index.aff": aff, "index.dic": dic }
const served = async (url: string) => ({ bytes: new Uint8Array(files[url.split("/").pop()!]) })

describe("spelling languages", () => {
  beforeEach(() => {
    vi.stubGlobal("indexedDB", memoryIndexedDb())
    useDictionaries.setState({ installed: [], downloading: {}, failed: null })
  })
  afterEach(() => vi.unstubAllGlobals())

  it("pick the dictionary for a region: its own, else its language's", () => {
    expect(dictionaryForLocale("nl-NL")).toBe("nl-NL")
    expect(dictionaryForLocale("nl-BE")).toBe("nl-NL")
    expect(dictionaryForLocale("en-IE")).toBe("en-US")
    expect(dictionaryForLocale("ja-JP")).toBeNull()
    expect(dictionaryForLocale("")).toBeNull()
  })

  it("are downloaded, checked, kept on the device and offered to Spelling", async () => {
    expect(await downloadDictionary("nl-NL", served)).toBe(true)
    expect(useDictionaries.getState()).toMatchObject({ installed: ["nl-NL"], downloading: {}, failed: null })
    expect((await storedDictionary("nl-NL"))!.dic.length).toBe(dic.length)
    expect(usableLanguages(useDictionaries.getState().installed).map((l) => l.id)).toEqual(["en-US", "nl-NL"])
    // After a restart, the kept ones are found again.
    useDictionaries.setState({ installed: [] })
    await loadInstalled()
    expect(useDictionaries.getState().installed).toEqual(["nl-NL"])
    await removeDictionary("nl-NL")
    expect(useDictionaries.getState().installed).toEqual([])
    expect(await storedDictionary("nl-NL")).toBeNull()
  })

  it("a file that isn't the pinned one, or no connection, keeps nothing and says why", async () => {
    const changed = async (url: string) => ({ bytes: new TextEncoder().encode(url) })
    expect(await downloadDictionary("de-DE", changed)).toBe(false)
    expect(useDictionaries.getState().failed).toEqual({ id: "de-DE", message: DICTIONARY_MESSAGES.changed })
    expect(await downloadDictionary("de-DE", () => Promise.reject(new Error("offline")))).toBe(false)
    expect(useDictionaries.getState().failed!.message).toBe(DICTIONARY_MESSAGES.offline)
    expect(useDictionaries.getState().installed).toEqual([])
    expect(await storedDictionary("de-DE")).toBeNull()
  })
})
