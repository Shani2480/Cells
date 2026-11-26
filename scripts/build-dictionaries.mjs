// Builds app/src/proofing/dictionaries-data.ts: the spelling dictionaries Cells offers as downloads (wooorm's
// dictionary-* packages, Hunspell files), each pinned to a version on jsDelivr with the files' size and SHA-256, so
// a download that doesn't match is refused. Run: node scripts/build-dictionaries.mjs
// Left out: Italian and Portuguese (both), whose files nspell can't read ("Too many properties to enumerate").
import { createHash } from "node:crypto"
import { writeFileSync } from "node:fs"

const LANGS = [
  ["en-GB", "en-gb", "English (United Kingdom)"],
  ["en-AU", "en-au", "English (Australia)"],
  ["en-CA", "en-ca", "English (Canada)"],
  ["nl-NL", "nl", "Dutch"],
  ["de-DE", "de", "German"],
  ["fr-FR", "fr", "French"],
  ["es-ES", "es", "Spanish"],
  ["sv-SE", "sv", "Swedish"],
  ["da-DK", "da", "Danish"],
  ["nb-NO", "nb", "Norwegian Bokmål"],
  ["pl-PL", "pl", "Polish"],
  ["ru-RU", "ru", "Russian"],
  ["cs-CZ", "cs", "Czech"],
  ["tr-TR", "tr", "Turkish"],
]

const registry = async (name) => (await fetch(`https://registry.npmjs.org/${name}/latest`)).json()
const out = []
for (const [id, pkg, label] of LANGS) {
  const meta = await registry(`dictionary-${pkg}`)
  const base = `https://cdn.jsdelivr.net/npm/dictionary-${pkg}@${meta.version}/`
  const files = {}
  for (const f of ["aff", "dic"]) {
    const bytes = new Uint8Array(await (await fetch(`${base}index.${f}`)).arrayBuffer())
    files[f] = { size: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex") }
  }
  out.push({ id, label, license: meta.license, base, ...files })
  console.log(id, meta.version, files.dic.size)
}
const ts =
  `// Made by scripts/build-dictionaries.mjs — don't edit by hand.\n` +
  `import type { DictionaryEntry } from "./dictionaries"\n\n` +
  `export const DOWNLOADABLE_DICTIONARIES: DictionaryEntry[] = ${JSON.stringify(out, null, 2)}\n`
writeFileSync(new URL("../app/src/proofing/dictionaries-data.ts", import.meta.url), ts)
