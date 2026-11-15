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
})
