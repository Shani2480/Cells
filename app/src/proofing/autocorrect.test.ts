import { describe, expect, it } from "vitest"
import { autoCorrectAt, autoCorrectEntries, DEFAULT_AUTOCORRECT, type AutoCorrectSettings } from "./autocorrect"

const fix = (text: string, s: AutoCorrectSettings = DEFAULT_AUTOCORRECT) =>
  autoCorrectAt(text, text.length, s)?.text ?? text

describe("AutoCorrect", () => {
  it("replaces entries of the list as they're finished, keeping the case typed", () => {
    expect(fix("teh")).toBe("the")
    expect(fix("Teh")).toBe("The")
    expect(fix("TEH")).toBe("THE")
    expect(fix("tEh")).toBe("tEh")
    expect(fix("Total (c)")).toBe("Total ©")
    expect(fix("A -->")).toBe("A →")
    expect(fix("(teh")).toBe("(the")
    expect(fix("i")).toBe("I")
  })

  it("puts right two initial capitals, Caps Lock words and day names", () => {
    expect(fix("THe")).toBe("The")
    expect(fix("IDs")).toBe("IDs")
    expect(fix("NASA")).toBe("NASA")
    expect(fix("hELLO")).toBe("Hello")
    expect(fix("due monday")).toBe("due Monday")
  })

  it("capitalizes a sentence's first word, but not a cell's", () => {
    expect(fix("done. next")).toBe("done. Next")
    expect(fix("next")).toBe("next")
    expect(fix("see e.g. this")).toBe("see e.g. this")
  })

  it("leaves formulas, and does what the options and the list say", () => {
    expect(fix("=teh")).toBe("=teh")
    expect(fix("-teh")).toBe("-teh")
    const off: AutoCorrectSettings = { ...DEFAULT_AUTOCORRECT, replaceText: false, dayNames: false }
    expect(fix("teh monday", off)).toBe("teh monday")
    const own: AutoCorrectSettings = { ...DEFAULT_AUTOCORRECT, added: [["ytd", "year to date"]], removed: ["teh"] }
    expect(fix("ytd", own)).toBe("year to date")
    expect(fix("teh", own)).toBe("teh")
    expect(autoCorrectEntries(own).get("adn")).toBe("and")
  })

  it("corrects only the word that ends where the separator went", () => {
    // "teh teh|" with the caret after the first word: only it changes.
    expect(autoCorrectAt("teh teh", 3, DEFAULT_AUTOCORRECT)).toEqual({ text: "the teh", delta: 0 })
    expect(autoCorrectAt("all fine", 8, DEFAULT_AUTOCORRECT)).toBeNull()
  })
})
