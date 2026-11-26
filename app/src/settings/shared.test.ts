import { describe, expect, it } from "vitest"
import { DEFAULT_SETTINGS, detectSettings, hasSharedValues, mergeShared, nameFromAccount, parseShared } from "./shared"

const inkspire = JSON.stringify({
  version: 1,
  user: { name: "Alex Demo", initials: "AD", avatar: "kept" },
  regional: { locale: "en-GB", dateFormat: "dd/MM/yyyy", decimalSeparator: ".", groupSeparator: "," },
  appearance: { theme: "dark", accent: "green" },
  apps: { inkspire: { setupDone: true, ribbon: "simplified" } },
  ai: { provider: "builtin" },
})

describe("the settings file shared with Inkspire", () => {
  it("reads what Inkspire wrote; Cells' own setup is still to do", () => {
    expect(parseShared(inkspire)).toEqual({
      name: "Alex Demo",
      initials: "AD",
      locale: "en-GB",
      dateFormat: "dd/MM/yyyy",
      decimalSeparator: ".",
      groupSeparator: ",",
      theme: "dark",
      setupDone: false,
      aiEnabled: false,
      ocrEnabled: false,
      autoFlashFill: true,
      aiProvider: "builtin",
      aiModels: {},
      aiAddresses: {},
      dictionaryWords: [],
      hintsEnabled: true,
      hintsDismissed: [],
    })
    expect(hasSharedValues(inkspire)).toBe(true)
  })

  it("reads a missing, broken or odd file as the defaults", () => {
    expect(parseShared(null)).toEqual(DEFAULT_SETTINGS)
    expect(parseShared("{not json")).toEqual(DEFAULT_SETTINGS)
    expect(parseShared("[1,2]")).toEqual(DEFAULT_SETTINGS)
    expect(parseShared(JSON.stringify({ user: { name: 7 }, appearance: { theme: "pink" } }))).toEqual(DEFAULT_SETTINGS)
    expect(hasSharedValues(null)).toBe(false)
  })

  it("writes its own values in and keeps every other key", () => {
    const merged = JSON.parse(mergeShared(inkspire, { name: "Sam Example", theme: "light", setupDone: true }))
    expect(merged.user).toEqual({ name: "Sam Example", initials: "AD", avatar: "kept" })
    expect(merged.appearance).toEqual({ theme: "light", accent: "green" })
    expect(merged.apps).toEqual({
      inkspire: { setupDone: true, ribbon: "simplified" },
      cells: { setupDone: true },
    })
    expect(merged.ai).toEqual({ provider: "builtin" })
    expect(merged.version).toBe(1)
  })

  it("keeps feature hints on by default, and Cells' dismissed hints and switch under apps.cells", () => {
    expect(parseShared(null)).toMatchObject({ hintsEnabled: true, hintsDismissed: [] })
    const written = mergeShared('{"apps": {"inkspire": {"x": 1}}}', {
      hintsEnabled: false,
      hintsDismissed: ["ai-tools", "flash-fill"],
    })
    expect(JSON.parse(written).apps).toEqual({
      inkspire: { x: 1 },
      cells: { hints: false, hintsDismissed: ["ai-tools", "flash-fill"] },
    })
    expect(parseShared(written)).toMatchObject({ hintsEnabled: false, hintsDismissed: ["ai-tools", "flash-fill"] })
    // Show hints again empties the list (it isn't merged like added words).
    expect(parseShared(mergeShared(written, { hintsDismissed: [] })).hintsDismissed).toEqual([])
  })

  it("keeps a newer version number", () => {
    expect(JSON.parse(mergeShared(JSON.stringify({ version: 3 }), {})).version).toBe(3)
    expect(JSON.parse(mergeShared(null, DEFAULT_SETTINGS))).toMatchObject({
      version: 1,
      apps: { cells: { setupDone: false } },
    })
  })
})

describe("guessing from the system (as Inkspire does)", () => {
  it("takes the name and region from the session", () => {
    const s = detectSettings({ user: "alex", full_name: "Alex Demo", lang: "en_GB.UTF-8" })
    expect(s).toMatchObject({ name: "Alex Demo", initials: "AD", locale: "en-GB", dateFormat: "dd/MM/yyyy" })
    expect(s).toMatchObject({ decimalSeparator: ".", groupSeparator: "," })
  })

  it("lets LC_TIME and LC_NUMERIC set the date and number formats", () => {
    const s = detectSettings({ lang: "en_US.UTF-8", lc_time: "de_DE.UTF-8", lc_numeric: "de_DE.UTF-8" })
    expect(s).toMatchObject({ locale: "en-US", dateFormat: "dd.MM.yyyy", decimalSeparator: ",", groupSeparator: "." })
  })

  it("falls back to the browser's locale, and leaves the name blank rather than guess", () => {
    expect(detectSettings({ lang: "C", intlLocale: "nl-NL" })).toMatchObject({ locale: "nl-NL", name: "" })
    expect(nameFromAccount("jane_doe")).toBe("Jane Doe")
    expect(nameFromAccount("alex99")).toBe("")
  })

  it("keeps the AI choice with Inkspire's, and models of other providers", () => {
    const source = JSON.stringify({
      ai: {
        provider: "ollama",
        models: { ollama: "llama3.2", openai: "gpt-5" },
        addresses: { lmstudio: "http://127.0.0.1:1234/v1" },
      },
      apps: { inkspire: { setupDone: true } },
    })
    const read = parseShared(source)
    expect(read).toMatchObject({
      aiEnabled: false,
      aiProvider: "ollama",
      aiModels: { ollama: "llama3.2", openai: "gpt-5" },
    })
    const merged = JSON.parse(
      mergeShared(source, { aiEnabled: true, aiProvider: "lmstudio", aiModels: { lmstudio: "qwen3-8b" } }),
    )
    expect(merged.ai).toEqual({
      provider: "lmstudio",
      models: { ollama: "llama3.2", openai: "gpt-5", lmstudio: "qwen3-8b" },
      addresses: { lmstudio: "http://127.0.0.1:1234/v1" },
    })
    expect(merged.apps).toEqual({ inkspire: { setupDone: true }, cells: { ai: true } })
  })
})

describe("dictionary words", () => {
  it("adds to the file's list, keeping words the other app added", () => {
    const file = JSON.stringify({ proofing: { customWords: ["Zorbly"] } })
    const merged = mergeShared(file, { dictionaryWords: ["Flurbo"] })
    expect(parseShared(merged).dictionaryWords).toEqual(["Zorbly", "Flurbo"])
  })
})
