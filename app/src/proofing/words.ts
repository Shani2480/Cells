/**
 * Words in a cell's text, for Spelling: runs of letters (with apostrophes inside, as in "don't"), each with where
 * it is. Excel's default options leave out words in UPPERCASE, words with numbers, and internet and file addresses.
 */

export type WordAt = { word: string; start: number; end: number }

export type SpellOptions = { ignoreUppercase: boolean; ignoreNumbers: boolean; ignoreAddresses: boolean }

export const DEFAULT_SPELL_OPTIONS: SpellOptions = { ignoreUppercase: true, ignoreNumbers: true, ignoreAddresses: true }

/** Internet addresses, e-mail addresses and file paths, which aren't checked. */
const ADDRESS = /(?:[a-z][a-z0-9+.-]*:\/\/|www\.)\S+|\S+@\S+\.\S+|(?:[a-z]:)?[\\/]\S+[\\/]\S*/gi

const TOKEN = /[\p{L}\p{N}]+(?:['’][\p{L}\p{N}]+)*/gu

export function wordsIn(text: string, options: SpellOptions = DEFAULT_SPELL_OPTIONS): WordAt[] {
  const skip: [number, number][] = []
  if (options.ignoreAddresses) for (const m of text.matchAll(ADDRESS)) skip.push([m.index, m.index + m[0].length])
  const out: WordAt[] = []
  for (const m of text.matchAll(TOKEN)) {
    const start = m.index
    const end = start + m[0].length
    const word = m[0]
    if (skip.some(([a, b]) => start < b && end > a)) continue
    if (!/\p{L}/u.test(word)) continue
    if (options.ignoreNumbers && /\p{N}/u.test(word)) continue
    // Excel ignores words in capitals (acronyms), and single letters.
    if (options.ignoreUppercase && word.length > 1 && word === word.toUpperCase() && word !== word.toLowerCase())
      continue
    if ([...word].length < 2) continue
    out.push({ word, start, end })
  }
  return out
}

/** A replacement takes the case pattern of the word it replaces (Excel: "teh" → "the", "Teh" → "The"). */
export function matchCase(original: string, replacement: string): string {
  if (original.length > 1 && original === original.toUpperCase()) return replacement.toUpperCase()
  if (original[0] && original[0] === original[0].toUpperCase() && original[0] !== original[0].toLowerCase())
    return replacement[0].toUpperCase() + replacement.slice(1)
  return replacement
}
