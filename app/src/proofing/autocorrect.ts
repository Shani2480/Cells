import { matchCase } from "./words"

/**
 * AutoCorrect, as Office's (File › Options › Proofing › AutoCorrect Options): while text is typed into a cell, the
 * word just finished (by a space, punctuation or the end of the entry) is corrected. Replace text as you type
 * swaps entries of the list ("teh" → "the", "(c)" → "©"), keeping the case typed; TWo INitial CApitals, the names of
 * days and words typed with Caps Lock on are put right; the first letter after a sentence's end is capitalized
 * (Excel doesn't capitalize a cell's first word). Formulas aren't touched.
 */

export type AutoCorrectSettings = {
  twoInitialCaps: boolean
  sentenceCaps: boolean
  dayNames: boolean
  capsLock: boolean
  replaceText: boolean
  /** Entries the person added or changed (replacing the built-in one of the same name). */
  added: [string, string][]
  /** Built-in entries the person deleted. */
  removed: string[]
}

export const DEFAULT_AUTOCORRECT: AutoCorrectSettings = {
  twoInitialCaps: true,
  sentenceCaps: true,
  dayNames: true,
  capsLock: true,
  replaceText: true,
  added: [],
  removed: [],
}

/** Office's built-in English entries that matter in a spreadsheet: symbols and common slips. */
export const BUILT_IN_ENTRIES: [string, string][] = [
  ["(c)", "©"],
  ["(e)", "€"],
  ["(r)", "®"],
  ["(tm)", "™"],
  ["...", "…"],
  ["-->", "→"],
  ["<--", "←"],
  ["<->", "↔"],
  ["==>", "⇒"],
  ["<==", "⇐"],
  ["<=>", "⇔"],
  ["abbout", "about"],
  ["abotu", "about"],
  ["accomodate", "accommodate"],
  ["accross", "across"],
  ["acheive", "achieve"],
  ["acn", "can"],
  ["adn", "and"],
  ["agian", "again"],
  ["alot", "a lot"],
  ["alwasy", "always"],
  ["amde", "made"],
  ["andthe", "and the"],
  ["arent", "aren't"],
  ["artical", "article"],
  ["becasue", "because"],
  ["becuase", "because"],
  ["beggining", "beginning"],
  ["beleive", "believe"],
  ["bianry", "binary"],
  ["busines", "business"],
  ["calender", "calendar"],
  ["cemetary", "cemetery"],
  ["certian", "certain"],
  ["comming", "coming"],
  ["commitee", "committee"],
  ["companys", "companies"],
  ["comapny", "company"],
  ["couldnt", "couldn't"],
  ["definately", "definitely"],
  ["didnt", "didn't"],
  ["doesnt", "doesn't"],
  ["dont", "don't"],
  ["eahc", "each"],
  ["embarass", "embarrass"],
  ["enviroment", "environment"],
  ["esle", "else"],
  ["excercise", "exercise"],
  ["existance", "existence"],
  ["familar", "familiar"],
  ["febuary", "February"],
  ["foriegn", "foreign"],
  ["freind", "friend"],
  ["fromthe", "from the"],
  ["goverment", "government"],
  ["gaurd", "guard"],
  ["hadnt", "hadn't"],
  ["hasnt", "hasn't"],
  ["havent", "haven't"],
  ["hte", "the"],
  ["i", "I"],
  ["immediatly", "immediately"],
  ["independant", "independent"],
  ["isnt", "isn't"],
  ["knowlege", "knowledge"],
  ["libary", "library"],
  ["maintenence", "maintenance"],
  ["managment", "management"],
  ["mispell", "misspell"],
  ["neccessary", "necessary"],
  ["necesary", "necessary"],
  ["noticable", "noticeable"],
  ["occassion", "occasion"],
  ["occured", "occurred"],
  ["occurence", "occurrence"],
  ["ofthe", "of the"],
  ["peice", "piece"],
  ["posession", "possession"],
  ["prefered", "preferred"],
  ["presense", "presence"],
  ["probelm", "problem"],
  ["publically", "publicly"],
  ["recieve", "receive"],
  ["recieved", "received"],
  ["reccomend", "recommend"],
  ["recomend", "recommend"],
  ["refered", "referred"],
  ["relevent", "relevant"],
  ["resturant", "restaurant"],
  ["seperate", "separate"],
  ["shouldnt", "shouldn't"],
  ["similiar", "similar"],
  ["succesful", "successful"],
  ["sucessful", "successful"],
  ["teh", "the"],
  ["tehy", "they"],
  ["thier", "their"],
  ["thsi", "this"],
  ["tommorow", "tomorrow"],
  ["tomorow", "tomorrow"],
  ["truely", "truly"],
  ["untill", "until"],
  ["wasnt", "wasn't"],
  ["wednesay", "Wednesday"],
  ["werent", "weren't"],
  ["wich", "which"],
  ["wierd", "weird"],
  ["wiht", "with"],
  ["wouldnt", "wouldn't"],
  ["yera", "year"],
  ["yoru", "your"],
]

/** Words that keep two capitals at the start (Office's INitial CAps exceptions): plurals of abbreviations. */
const TWO_CAPS_EXCEPTIONS = new Set(["IDs", "PCs", "CDs", "TVs", "OKs", "URLs", "PDFs", "APIs"])

const DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

/** The list in use: the built-in entries less those deleted, with the person's own over them. */
export function autoCorrectEntries(s: AutoCorrectSettings): Map<string, string> {
  const map = new Map<string, string>()
  for (const [from, to] of BUILT_IN_ENTRIES) if (!s.removed.includes(from)) map.set(from, to)
  for (const [from, to] of s.added) map.set(from, to)
  return map
}

/** What ends a word as it's typed: white space and punctuation (an apostrophe stays inside: "don't"). */
export const isWordEnd = (ch: string) => /[\s.,;:!?)\]}"”]/.test(ch)

/** An entry for what was typed: exactly, or (for an entry in lower case) typed capitalized or in capitals. */
function lookup(entries: Map<string, string>, typed: string): string | null {
  const exact = entries.get(typed)
  if (exact !== undefined) return exact
  const lower = typed.toLowerCase()
  if (lower === typed) return null
  const to = entries.get(lower)
  if (to === undefined) return null
  const capitalized = typed[0] + typed.slice(1).toLowerCase() === typed
  if (typed !== typed.toUpperCase() && !capitalized) return null
  return matchCase(typed, to)
}

/** A word's own fixes: two initial capitals, Caps Lock, day names. Null: nothing to fix. */
function fixWord(word: string, s: AutoCorrectSettings): string | null {
  if (!/^\p{L}+$/u.test(word)) return null
  const lower = word.toLowerCase()
  const upper = word.toUpperCase()
  if (
    s.capsLock &&
    word.length > 2 &&
    word[0] === lower[0] &&
    word[0] !== upper[0] &&
    word.slice(1) === upper.slice(1) &&
    upper.slice(1) !== lower.slice(1)
  )
    return word[0].toUpperCase() + word.slice(1).toLowerCase()
  if (
    s.twoInitialCaps &&
    word.length > 2 &&
    word.slice(0, 2) === upper.slice(0, 2) &&
    word.slice(2) === lower.slice(2) &&
    word.slice(2) !== upper.slice(2) &&
    word.slice(0, 2) !== lower.slice(0, 2) &&
    !TWO_CAPS_EXCEPTIONS.has(word)
  )
    return word[0] + word.slice(1).toLowerCase()
  if (s.dayNames && DAYS.includes(word)) return word[0].toUpperCase() + word.slice(1)
  return null
}

/**
 * The text with the word that ends at `end` corrected (the text before `end`, back to white space), and how many
 * characters it grew by; null when nothing changes. `end` is where the separator was typed (or the text's end).
 */
export function autoCorrectAt(
  text: string,
  end: number,
  s: AutoCorrectSettings,
  entries = autoCorrectEntries(s),
): { text: string; delta: number } | null {
  // A formula (Excel also reads an entry starting with + or - as one).
  if (/^[=+-]/.test(text)) return null
  let start = end
  while (start > 0 && !/\s/.test(text[start - 1])) start--
  const token = text.slice(start, end)
  if (!token) return null
  let replaced: string | null = null
  let from = start
  let to = end
  // Replace text: the whole token ("(c)", "-->"), or the word inside its punctuation ("teh," "(teh").
  if (s.replaceText) {
    replaced = lookup(entries, token)
    if (replaced === null) {
      const m = /^([("'“‘[]*)(.*?)([)"'”’\]]*)$/u.exec(token)!
      const inner = m[2]
      const hit = inner ? lookup(entries, inner) : null
      if (hit !== null) {
        replaced = hit
        from = start + m[1].length
        to = from + inner.length
      }
    }
  }
  if (replaced === null) {
    const m = /^([^\p{L}]*)(\p{L}+)([^\p{L}]*)$/u.exec(token)
    if (m) {
      const fixed = fixWord(m[2], s)
      if (fixed !== null) {
        replaced = fixed
        from = start + m[1].length
        to = from + m[2].length
      }
    }
  }
  // The first letter of a sentence (after . ! or ? and a space, not a cell's first word).
  let out = replaced === null ? text : text.slice(0, from) + replaced + text.slice(to)
  if (s.sentenceCaps) {
    const before = out.slice(0, from)
    const word = out.slice(from, from + (replaced ?? text.slice(from, to)).length)
    if (/[.!?]\s+$/.test(before) && /^\p{Ll}/u.test(word) && !/\.\S+\.\s+$/.test(before)) {
      out = before + word[0].toUpperCase() + out.slice(from + 1)
    }
  }
  if (out === text) return null
  return { text: out, delta: out.length - text.length }
}

let active: { settings: AutoCorrectSettings; entries: Map<string, string> } | null = {
  settings: DEFAULT_AUTOCORRECT,
  entries: autoCorrectEntries(DEFAULT_AUTOCORRECT),
}

/** The person's AutoCorrect options (from the shared settings); null turns it off (tests that type exact text). */
export function setAutoCorrect(settings: AutoCorrectSettings | null) {
  active = settings ? { settings, entries: autoCorrectEntries(settings) } : null
}

/**
 * What typing did to an entry, corrected: when the one character just typed at `caret - 1` ends a word, the word
 * before it (or, for punctuation that completes an entry such as "(c)", the token with it). Null: nothing to do.
 */
export function autoCorrectTyped(text: string, caret: number): { text: string; caret: number } | null {
  if (!active || caret < 1 || !isWordEnd(text[caret - 1])) return null
  const { settings, entries } = active
  const word = autoCorrectAt(text, caret - 1, settings, entries)
  const fixed = word ?? (/\s/.test(text[caret - 1]) ? null : autoCorrectAt(text, caret, settings, entries))
  return fixed ? { text: fixed.text, caret: caret + fixed.delta } : null
}

/** An entry as it goes into its cell: its last word corrected too. */
export function autoCorrectEntry(text: string): string {
  if (!active) return text
  return autoCorrectAt(text, text.length, active.settings, active.entries)?.text ?? text
}
