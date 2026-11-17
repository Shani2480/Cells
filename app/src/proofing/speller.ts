import nspell from "nspell"
import affUrl from "dictionary-en-files/index.aff?url"
import dicUrl from "dictionary-en-files/index.dic?url"

/**
 * Spelling dictionaries: Hunspell word lists that ship with Cells (English, from SCOWL via the dictionary-en
 * package), read by nspell. Nothing is looked up online. Words the person added are shared with Inkspire.
 */

export type Speller = {
  correct(word: string): boolean
  suggest(word: string): string[]
  add(word: string): void
}

/** The dictionary languages Cells has (more can come as downloads). */
export const SPELL_LANGUAGES = [{ id: "en-US", label: "English (United States)" }] as const

export function spellerFrom(aff: string, dic: string, words: readonly string[] = []): Speller {
  const s = nspell({ aff, dic })
  for (const w of words) s.add(w)
  return {
    correct: (word) => s.correct(word) || s.correct(word.replace(/’/g, "'")),
    suggest: (word) => s.suggest(word).slice(0, 8),
    add: (word) => void s.add(word),
  }
}

let loading: Promise<Speller> | null = null

/** The English dictionary, read once (the files come with the app). */
export function loadSpeller(words: readonly string[]): Promise<Speller> {
  loading ??= Promise.all([fetch(affUrl).then((r) => r.text()), fetch(dicUrl).then((r) => r.text())]).then(
    ([aff, dic]) => spellerFrom(aff, dic, words),
  )
  return loading
}
