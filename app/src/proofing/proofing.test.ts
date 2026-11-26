import { readFileSync } from "node:fs"
import { join } from "node:path"
import { gunzipSync } from "node:zlib"
import { describe, expect, it } from "vitest"
import { spellerFrom } from "./speller"
import { baseForms, parseThesaurus } from "./thesaurus"
import { matchCase, wordsIn } from "./words"

const root = join(__dirname, "../../..")
const speller = spellerFrom(
  readFileSync(join(root, "node_modules/dictionary-en/index.aff"), "utf8"),
  readFileSync(join(root, "node_modules/dictionary-en/index.dic"), "utf8"),
  ["Cells"],
)
const thesaurus = parseThesaurus(
  gunzipSync(readFileSync(join(root, "app/public/proofing/en/thesaurus.txt.gz"))).toString("utf8"),
)

describe("words to check", () => {
  it("skips capitals, words with numbers, addresses and single letters", () => {
    const words = wordsIn("Teh NASA report Q3 at www.example.com or alex@example.com, a don't test2")
    expect(words.map((w) => w.word)).toEqual(["Teh", "report", "at", "or", "don't"])
    expect(words[0]).toEqual({ word: "Teh", start: 0, end: 3 })
  })

  it("keeps the case of the word it replaces", () => {
    expect(matchCase("Teh", "the")).toBe("The")
    expect(matchCase("TEH", "the")).toBe("THE")
    expect(matchCase("teh", "the")).toBe("the")
  })
})

describe("the English dictionary", () => {
  it("knows words, suggests for misspellings, and takes added words", async () => {
    expect((await speller.check(["receive"]))[0]).toBe(true)
    expect((await speller.check(["recieve"]))[0]).toBe(false)
    expect(await speller.suggest("recieve")).toContain("receive")
    expect((await speller.check(["Cells"]))[0]).toBe(true)
    expect((await speller.check(["Zorbly"]))[0]).toBe(false)
    speller.add("Zorbly")
    expect((await speller.check(["Zorbly"]))[0]).toBe(true)
  })
})

describe("the thesaurus", () => {
  it("lists a word's meanings, most common first, with synonyms and opposites", () => {
    const happy = thesaurus.lookup("happy")
    expect(happy[0].pos).toBe("adj.")
    expect(happy.flatMap((m) => m.synonyms)).toContain("glad")
    expect(happy.flatMap((m) => m.antonyms)).toContain("unhappy")
  })

  it("finds the base form of an inflected word", () => {
    expect(baseForms("cities")).toContain("city")
    expect(thesaurus.lookup("running").length).toBeGreaterThan(0)
    expect(thesaurus.lookup("qqqq")).toEqual([])
  })
})
