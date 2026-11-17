// Builds Cells' English thesaurus (app/public/proofing/en/thesaurus.txt.gz) from WordNet's database files.
//
//   node scripts/build-thesaurus.mjs <wordnet dict folder>
//
// The folder is WordNet 3.x's `dict` (data.noun, data.verb, data.adj, data.adv and their index files), e.g. from
// the wordnet-db npm package. WordNet is © Princeton University; its licence goes alongside the file.
//
// The output, one line each:
//   M <pos> <words, comma-separated> <related meanings> <opposite meanings>   (fields tab-separated; meanings are
//                                                                              line numbers among the M lines)
//   W <word> <meanings in WordNet's order, most common first>                  (only words with several meanings)
import { readFileSync, writeFileSync, mkdirSync, copyFileSync } from "node:fs"
import path from "node:path"
import { gzipSync } from "node:zlib"

const dict = process.argv[2]
if (!dict) {
  console.error("usage: node scripts/build-thesaurus.mjs <wordnet dict folder>")
  process.exit(1)
}
const POS = { noun: "n", verb: "v", adj: "a", adv: "r" }
const clean = (w) => w.replace(/\((a|p|ip)\)$/, "").replace(/_/g, " ")

/** Synsets by "pos:offset". */
const synsets = new Map()
for (const [file, pos] of Object.entries(POS)) {
  for (const line of readFileSync(path.join(dict, `data.${file}`), "latin1").split("\n")) {
    if (!line || line.startsWith("  ")) continue
    const head = line.split(" | ")[0].split(" ")
    const offset = head[0]
    const count = parseInt(head[3], 16)
    const words = []
    let i = 4
    for (let k = 0; k < count; k++, i += 2) words.push(clean(head[i]))
    const pointers = Number(head[i++])
    const related = []
    const opposite = []
    for (let k = 0; k < pointers; k++, i += 4) {
      const [symbol, target, targetPos] = [head[i], head[i + 1], head[i + 2]]
      const key = `${targetPos === "s" ? "a" : targetPos}:${target}`
      // Similar adjectives and broader terms are offered as related words; antonyms as opposites.
      if (symbol === "&" || symbol === "@") related.push(key)
      else if (symbol === "!") opposite.push(key)
    }
    synsets.set(`${pos}:${offset}`, { pos, words, related, opposite })
  }
}

const keys = [...synsets.keys()]
const indexOf = new Map(keys.map((k, i) => [k, i]))
const lines = keys.map((k) => {
  const s = synsets.get(k)
  const ids = (list) => [...new Set(list.map((x) => indexOf.get(x)).filter((x) => x !== undefined))].join(",")
  return `M\t${s.pos}\t${[...new Set(s.words)].join(",")}\t${ids(s.related)}\t${ids(s.opposite)}`
})

// Sense order for words with several meanings (WordNet's index files list the most common first).
const senses = new Map()
for (const [file, pos] of Object.entries(POS)) {
  for (const line of readFileSync(path.join(dict, `index.${file}`), "latin1").split("\n")) {
    if (!line || line.startsWith("  ")) continue
    const f = line.trim().split(" ")
    const lemma = clean(f[0]).toLowerCase()
    const synsetCount = Number(f[2])
    const offsets = f.slice(f.length - synsetCount)
    const list = senses.get(lemma) ?? []
    for (const o of offsets) {
      const id = indexOf.get(`${pos === "a" ? "a" : pos}:${o}`)
      if (id !== undefined) list.push(id)
    }
    senses.set(lemma, list)
  }
}
for (const [lemma, list] of senses) if (list.length > 1) lines.push(`W\t${lemma}\t${list.join(",")}`)

const out = path.resolve(path.dirname(new URL(import.meta.url).pathname), "../app/public/proofing/en")
mkdirSync(out, { recursive: true })
const text = lines.join("\n") + "\n"
writeFileSync(path.join(out, "thesaurus.txt.gz"), gzipSync(Buffer.from(text, "utf8"), { level: 9 }))
const licence = ["LICENSE", "license"].map((n) => path.join(dict, "..", n)).find((p) => {
  try {
    readFileSync(p)
    return true
  } catch {
    return false
  }
})
if (licence) copyFileSync(licence, path.join(out, "WORDNET-LICENSE.txt"))
console.log(`${synsets.size} meanings, ${text.length} bytes, written to ${out}`)
