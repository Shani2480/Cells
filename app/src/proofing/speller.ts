import nspell from "nspell"
import affUrl from "dictionary-en-files/index.aff?url"
import dicUrl from "dictionary-en-files/index.dic?url"
import { BUILT_IN_LANGUAGE, languageLabel, storedDictionary } from "./dictionaries"

/**
 * Spelling dictionaries: Hunspell word lists read by nspell. English (United States) ships with Cells (SCOWL, via
 * the dictionary-en package); other languages are downloaded (./dictionaries). A large dictionary takes seconds to
 * read, so the app reads it in a worker and asks it about words in batches. Nothing is looked up online. Words the
 * person added are shared with Inkspire.
 */

export type Speller = {
  /** For each word, whether the dictionary knows it. */
  check(words: readonly string[]): Promise<boolean[]>
  suggest(word: string): Promise<string[]>
  add(word: string): void
}

/** A dictionary read here, in this thread (tests, and where workers aren't available). */
export function spellerFrom(aff: string, dic: string, words: readonly string[] = []): Speller {
  const s = nspell({ aff, dic })
  for (const w of words) s.add(w)
  const correct = (word: string) => s.correct(word) || s.correct(word.replace(/’/g, "'"))
  return {
    check: async (list) => list.map(correct),
    suggest: async (word) => s.suggest(word).slice(0, 8),
    add: (word) => void s.add(word),
  }
}

type Reply = { id: number; result: unknown; error?: string }

/** A dictionary read in a worker (./speller.worker.ts). */
function workerSpeller(aff: string, dic: string, words: readonly string[]): Promise<Speller> {
  const worker = new Worker(new URL("./speller.worker.ts", import.meta.url), { type: "module" })
  let next = 0
  const waiting = new Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void }>()
  worker.onmessage = (e: MessageEvent<Reply>) => {
    const w = waiting.get(e.data.id)
    waiting.delete(e.data.id)
    if (e.data.error) w?.reject(new Error(e.data.error))
    else w?.resolve(e.data.result)
  }
  const ask = <T>(message: Record<string, unknown>) =>
    new Promise<T>((resolve, reject) => {
      const id = next++
      waiting.set(id, { resolve: resolve as (v: unknown) => void, reject })
      worker.postMessage({ ...message, id })
    })
  return ask<void>({ kind: "load", aff, dic, words: [...words] }).then(() => ({
    check: (list) => ask<boolean[]>({ kind: "check", words: [...list] }),
    suggest: (word) => ask<string[]>({ kind: "suggest", word }),
    add: (word) => void worker.postMessage({ kind: "add", word, id: -1 }),
  }))
}

export const SPELLER_MESSAGES = {
  missing: (label: string) => `The ${label} dictionary isn't on this device. Download it in File › Options › Proofing.`,
}

/** The dictionary's files: the built-in one from the app, a downloaded one from the device's storage. */
async function dictionaryFiles(language: string): Promise<{ aff: string; dic: string }> {
  if (language === BUILT_IN_LANGUAGE.id)
    return Promise.all([fetch(affUrl).then((r) => r.text()), fetch(dicUrl).then((r) => r.text())]).then(
      ([aff, dic]) => ({ aff, dic }),
    )
  const stored = await storedDictionary(language)
  if (!stored) throw new Error(SPELLER_MESSAGES.missing(languageLabel(language) || language))
  return stored
}

let loaded: { language: string; speller: Promise<Speller> } | null = null

/** A language's dictionary, read once (the last one used stays ready; switching language reads the new one). */
export function loadSpeller(language: string, words: readonly string[]): Promise<Speller> {
  if (loaded?.language === language) return loaded.speller
  const speller = dictionaryFiles(language).then(({ aff, dic }) =>
    typeof Worker === "undefined" ? spellerFrom(aff, dic, words) : workerSpeller(aff, dic, words),
  )
  loaded = { language, speller }
  // A failed read is tried again next time.
  speller.catch(() => {
    if (loaded?.speller === speller) loaded = null
  })
  return speller
}
