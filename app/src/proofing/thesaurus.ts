/**
 * Review › Thesaurus: English synonyms from WordNet (© Princeton University), built into one file by
 * scripts/build-thesaurus.mjs and read here. A word's meanings come most common first, each with its part of speech,
 * its synonyms, related words and opposites.
 */

export type Meaning = {
  pos: "noun" | "verb" | "adj." | "adv."
  /** The meaning's name in the list: its first synonym other than the word looked up. */
  heading: string
  synonyms: string[]
  antonyms: string[]
}

const POS: Record<string, Meaning["pos"]> = { n: "noun", v: "verb", a: "adj.", r: "adv." }

type Synset = { pos: string; words: string[]; related: number[]; opposite: number[] }

export type Thesaurus = { lookup(word: string): Meaning[] }

const ids = (s: string | undefined) => (s ? s.split(",").map(Number) : [])

export function parseThesaurus(text: string): Thesaurus {
  const synsets: Synset[] = []
  const order = new Map<string, number[]>()
  const byWord = new Map<string, number[]>()
  for (const line of text.split("\n")) {
    const f = line.split("\t")
    if (f[0] === "M") {
      const id = synsets.length
      const words = f[2].split(",")
      synsets.push({ pos: f[1], words, related: ids(f[3]), opposite: ids(f[4]) })
      for (const w of words) {
        const key = w.toLowerCase()
        const list = byWord.get(key)
        if (list) list.push(id)
        else byWord.set(key, [id])
      }
    } else if (f[0] === "W") order.set(f[1], ids(f[2]))
  }

  const meaningsOf = (key: string): Meaning[] => {
    const list = order.get(key) ?? byWord.get(key) ?? []
    return list.map((id) => {
      const s = synsets[id]
      const others = s.words.filter((w) => w.toLowerCase() !== key)
      const related = s.related.flatMap((r) => synsets[r]?.words ?? []).filter((w) => w.toLowerCase() !== key)
      const synonyms = [...new Set([...others, ...related])].slice(0, 14)
      const antonyms = [...new Set(s.opposite.flatMap((r) => synsets[r]?.words ?? []))]
      return { pos: POS[s.pos] ?? "noun", heading: synonyms[0] ?? s.words[0], synonyms, antonyms }
    })
  }

  return {
    lookup(word) {
      const key = word.trim().toLowerCase()
      if (!key) return []
      for (const form of baseForms(key)) {
        const found = meaningsOf(form)
        if (found.length) return found
      }
      return []
    },
  }
}

/** The word, then what it may be an inflection of ("cities" → "city", "running" → "run"). */
export function baseForms(word: string): string[] {
  const out = [word]
  const add = (...words: string[]) => {
    for (const w of words) if (w.length > 1 && !out.includes(w)) out.push(w)
  }
  if (word.endsWith("ies")) add(word.slice(0, -3) + "y")
  if (word.endsWith("es")) add(word.slice(0, -2))
  if (word.endsWith("s")) add(word.slice(0, -1))
  if (word.endsWith("ied")) add(word.slice(0, -3) + "y")
  if (word.endsWith("ed")) add(word.slice(0, -2), word.slice(0, -1))
  if (word.endsWith("ing")) add(word.slice(0, -3), word.slice(0, -3) + "e")
  if (/(.)\1(ed|ing)$/.test(word)) add(word.replace(/(.)\1(ed|ing)$/, "$1"))
  if (word.endsWith("er")) add(word.slice(0, -2), word.slice(0, -1))
  if (word.endsWith("est")) add(word.slice(0, -3), word.slice(0, -2))
  return out
}

let loading: Promise<Thesaurus> | null = null

/** The thesaurus file, read and unpacked once (it comes with the app). */
export function loadThesaurus(): Promise<Thesaurus> {
  loading ??= fetch(`${import.meta.env.BASE_URL}proofing/en/thesaurus.txt.gz`)
    .then(async (r) => {
      if (!r.ok) throw new Error("The thesaurus couldn't be read.")
      const bytes = new Uint8Array(await r.arrayBuffer())
      // A web server may already have unpacked it (it sends .gz files as gzip-encoded text).
      if (bytes[0] !== 0x1f || bytes[1] !== 0x8b) return new TextDecoder().decode(bytes)
      return new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"))).text()
    })
    .then(parseThesaurus)
  loading.catch(() => (loading = null))
  return loading
}
