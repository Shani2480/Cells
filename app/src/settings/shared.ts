/**
 * The settings Cells shares with Inkspire, and how they're read from and merged into the shared file
 * (`docs/shared-settings.md`). Pure functions only; the store that uses them is `settings.ts`.
 *
 * The file is written by two apps, so Cells never writes it from memory: it reads the latest text, puts its own
 * values in at their places and leaves every other key (Inkspire's, a later version's) as it found it.
 */
import type { SystemFacts } from "@cells/platform"

export type ThemeChoice = "system" | "light" | "dark"

/**
 * The AI providers, as Inkspire names them. `lmstudio` and `compatible` both talk to an OpenAI-compatible endpoint
 * (LM Studio just starts from its usual address); `builtin` is the engine Inkspire runs itself, which Cells can't
 * run yet.
 */
export type AiProviderName = "builtin" | "ollama" | "lmstudio" | "compatible" | "openai" | "anthropic"
export const AI_PROVIDERS: AiProviderName[] = ["builtin", "ollama", "lmstudio", "compatible", "openai", "anthropic"]
export type AddressedProvider = "ollama" | "lmstudio" | "compatible"
export const DEFAULT_ADDRESSES: Record<AddressedProvider, string> = {
  ollama: "http://127.0.0.1:11434",
  lmstudio: "http://127.0.0.1:1234/v1",
  compatible: "",
}

export type SharedSettings = {
  /** The name notes and comments are signed with (File › Options › User name in Excel). */
  name: string
  initials: string
  /** A BCP 47 tag such as `en-GB`; blank: the system's. */
  locale: string
  /** A short date pattern such as `dd/MM/yyyy`; blank: the locale's. */
  dateFormat: string
  /** Blank: the locale's. */
  decimalSeparator: string
  groupSeparator: string
  theme: ThemeChoice
  /** Cells' own first-run setup has been finished or skipped (Inkspire keeps its own under `apps.inkspire`). */
  setupDone: boolean
  /** Cells' AI tools are turned on (opt-in, per app: `apps.cells.ai`). */
  aiEnabled: boolean
  /** Text recognition on this device reads scanned PDFs (opt-in: `apps.cells.ocr`). */
  ocrEnabled: boolean
  /** Flash Fill offers the rest of a column while an entry is typed (Excel's "Automatically Flash Fill", on: `apps.cells.autoFlashFill`). */
  autoFlashFill: boolean
  /** The provider both apps use (`ai.provider`); blank until one is chosen. */
  aiProvider: AiProviderName | ""
  /** A model per provider (`ai.models`); blank: the provider's default. */
  aiModels: Partial<Record<AiProviderName, string>>
  /** Addresses of local and compatible endpoints (`ai.addresses`); blank: the default. */
  aiAddresses: Partial<Record<AddressedProvider, string>>
  /** Words added to the dictionary (Spelling › Add to Dictionary), shared by both apps (`proofing.customWords`). */
  dictionaryWords: string[]
  /** The dictionary Spelling uses, as a BCP 47 tag (`proofing.language`); "en-US" comes with the apps. */
  spellLanguage: string
  /** File › Options › Proofing: words Spelling passes over (`proofing.ignoreUppercase` …); all on, as in Excel. */
  ignoreUppercase: boolean
  ignoreNumbers: boolean
  ignoreAddresses: boolean
  /** Hints that introduce Cells' special features are shown (on by default; `apps.cells.hints`). */
  hintsEnabled: boolean
  /** The hints already dismissed, by id (`apps.cells.hintsDismissed`); emptied by Show hints again. */
  hintsDismissed: string[]
}

export const SETTINGS_VERSION = 1

export const DEFAULT_SETTINGS: SharedSettings = {
  name: "",
  initials: "",
  locale: "",
  dateFormat: "",
  decimalSeparator: "",
  groupSeparator: "",
  theme: "system",
  setupDone: false,
  aiEnabled: false,
  ocrEnabled: false,
  autoFlashFill: true,
  aiProvider: "",
  aiModels: {},
  aiAddresses: {},
  dictionaryWords: [],
  spellLanguage: "en-US",
  ignoreUppercase: true,
  ignoreNumbers: true,
  ignoreAddresses: true,
  hintsEnabled: true,
  hintsDismissed: [],
}

type Json = Record<string, unknown>

const isObject = (value: unknown): value is Json => typeof value === "object" && value !== null && !Array.isArray(value)

function parseObject(text: string | null): Json {
  if (!text) return {}
  try {
    const value = JSON.parse(text) as unknown
    return isObject(value) ? value : {}
  } catch {
    return {}
  }
}

const section = (root: Json, key: string): Json => (isObject(root[key]) ? root[key] : {})
const text = (value: unknown) => (typeof value === "string" ? value : "")

/** The settings in a shared file's text; anything missing or of the wrong type is the default. */
export function parseShared(source: string | null): SharedSettings {
  const root = parseObject(source)
  const user = section(root, "user")
  const regional = section(root, "regional")
  const appearance = section(root, "appearance")
  const cells = section(section(root, "apps"), "cells")
  const ai = section(root, "ai")
  const proofing = section(root, "proofing")
  const theme = appearance.theme
  const strings = (value: unknown, keys: readonly string[]) =>
    Object.fromEntries(
      keys.filter((k) => isObject(value) && typeof value[k] === "string").map((k) => [k, (value as Json)[k]]),
    )
  return {
    name: text(user.name),
    initials: text(user.initials),
    locale: text(regional.locale),
    dateFormat: text(regional.dateFormat),
    decimalSeparator: text(regional.decimalSeparator),
    groupSeparator: text(regional.groupSeparator),
    theme: theme === "light" || theme === "dark" ? theme : "system",
    setupDone: cells.setupDone === true,
    aiEnabled: cells.ai === true,
    ocrEnabled: cells.ocr === true,
    autoFlashFill: cells.autoFlashFill !== false,
    aiProvider: AI_PROVIDERS.includes(ai.provider as AiProviderName) ? (ai.provider as AiProviderName) : "",
    aiModels: strings(ai.models, AI_PROVIDERS),
    aiAddresses: strings(ai.addresses, Object.keys(DEFAULT_ADDRESSES)),
    dictionaryWords: Array.isArray(proofing.customWords)
      ? proofing.customWords.filter((w): w is string => typeof w === "string" && w.trim() !== "")
      : [],
    spellLanguage: text(proofing.language) || "en-US",
    ignoreUppercase: proofing.ignoreUppercase !== false,
    ignoreNumbers: proofing.ignoreNumbers !== false,
    ignoreAddresses: proofing.ignoreAddresses !== false,
    hintsEnabled: cells.hints !== false,
    hintsDismissed: Array.isArray(cells.hintsDismissed)
      ? cells.hintsDismissed.filter((id): id is string => typeof id === "string")
      : [],
  }
}

/** Whether the file holds anything a person set (Inkspire's setup, or Cells' own), to say so in the setup. */
export function hasSharedValues(source: string | null): boolean {
  const s = parseShared(source)
  return Boolean(s.name || s.locale || s.dateFormat)
}

/** Where each setting lives in the file: `[section, key]`, or `["apps", "cells", key]` for Cells' own. */
const PLACES: Record<keyof SharedSettings, string[]> = {
  name: ["user", "name"],
  initials: ["user", "initials"],
  locale: ["regional", "locale"],
  dateFormat: ["regional", "dateFormat"],
  decimalSeparator: ["regional", "decimalSeparator"],
  groupSeparator: ["regional", "groupSeparator"],
  theme: ["appearance", "theme"],
  setupDone: ["apps", "cells", "setupDone"],
  aiEnabled: ["apps", "cells", "ai"],
  ocrEnabled: ["apps", "cells", "ocr"],
  autoFlashFill: ["apps", "cells", "autoFlashFill"],
  aiProvider: ["ai", "provider"],
  aiModels: ["ai", "models"],
  aiAddresses: ["ai", "addresses"],
  dictionaryWords: ["proofing", "customWords"],
  spellLanguage: ["proofing", "language"],
  ignoreUppercase: ["proofing", "ignoreUppercase"],
  ignoreNumbers: ["proofing", "ignoreNumbers"],
  ignoreAddresses: ["proofing", "ignoreAddresses"],
  hintsEnabled: ["apps", "cells", "hints"],
  hintsDismissed: ["apps", "cells", "hintsDismissed"],
}

/**
 * The file's text with the changed settings put in at their places. Only those: a value the other app changed since
 * Cells last read the file stays as that app wrote it. Every other key is kept too.
 */
export function mergeShared(source: string | null, changes: Partial<SharedSettings>): string {
  const root = parseObject(source)
  root.version = typeof root.version === "number" && root.version > SETTINGS_VERSION ? root.version : SETTINGS_VERSION
  for (const [key, value] of Object.entries(changes) as [keyof SharedSettings, unknown][]) {
    const path = PLACES[key]
    if (!path || value === undefined) continue
    let at = root
    for (const part of path.slice(0, -1)) {
      at[part] = isObject(at[part]) ? { ...at[part] } : {}
      at = at[part] as Json
    }
    const last = path[path.length - 1]
    // A map (models, addresses) is merged key by key, so the other app's entries for other providers stay.
    // Added words join the file's list, so a word the other app added meanwhile stays.
    if (key === "dictionaryWords" && Array.isArray(value)) {
      const had = Array.isArray(at[last]) ? (at[last] as unknown[]).filter((w) => typeof w === "string") : []
      at[last] = [...new Set([...had, ...value])]
      continue
    }
    at[last] = isObject(value) ? { ...(isObject(at[last]) ? at[last] : {}), ...value } : value
  }
  return JSON.stringify(root, null, 2) + "\n"
}

// What follows is taken from Inkspire (src/document/user-settings.ts), so both apps guess the same from the system.

/** `en_GB.UTF-8` → `en-GB`; `C`, `POSIX` and empty give "". */
export function localeFromEnvironment(value: string | undefined): string {
  const base = (value ?? "").trim().split(/[.@]/)[0]
  if (!base || base === "C" || base === "POSIX") return ""
  try {
    return Intl.getCanonicalLocales(base.replace(/_/g, "-"))[0] ?? ""
  } catch {
    return ""
  }
}

export function initialsOf(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 3)
    .map((word) => Array.from(word)[0]?.toUpperCase() ?? "")
    .join("")
}

/** The short date pattern a locale uses, such as `dd/MM/yyyy`; "" when it cannot be worked out. */
export function datePatternFor(locale: string): string {
  if (!locale) return ""
  try {
    const parts = new Intl.DateTimeFormat(locale, { day: "2-digit", month: "2-digit", year: "numeric" }).formatToParts(
      new Date(2031, 10, 22),
    )
    return parts
      .map((part) =>
        part.type === "day" ? "dd" : part.type === "month" ? "MM" : part.type === "year" ? "yyyy" : part.value,
      )
      .join("")
  } catch {
    return ""
  }
}

export function separatorsFor(locale: string): { decimal: string; group: string } {
  if (!locale) return { decimal: "", group: "" }
  try {
    const parts = new Intl.NumberFormat(locale).formatToParts(1234567.5)
    return {
      decimal: parts.find((part) => part.type === "decimal")?.value ?? "",
      group: parts.find((part) => part.type === "group")?.value ?? "",
    }
  } catch {
    return { decimal: "", group: "" }
  }
}

/**
 * A name read from an account name such as `john.smith`, `jane_doe` or `MariaGarcia`: only when it falls apart into
 * two or more plain words. `alex99` or `admin` say nothing about a person, so they give "".
 */
export function nameFromAccount(account: string): string {
  const words = account
    .trim()
    .replace(/([\p{Ll}])(\p{Lu})/gu, "$1 $2")
    .split(/[\s._-]+/)
    .filter(Boolean)
  if (words.length < 2 || !words.every((word) => /^\p{L}{2,}$/u.test(word))) return ""
  return words.map((word) => word[0]!.toUpperCase() + word.slice(1).toLowerCase()).join(" ")
}

/**
 * The setup's starting values from the system. The environment wins (`LC_ALL`, then `LC_TIME`/`LC_NUMERIC` for the
 * date and number formats, then `LANG`), the browser's locale fills in only when it says nothing, and what isn't
 * known stays blank.
 */
export function detectSettings(facts: SystemFacts): SharedSettings {
  const general = localeFromEnvironment(facts.lc_all) || localeFromEnvironment(facts.lang)
  const locale = general || localeFromEnvironment(facts.intlLocale?.replace("-", "_"))
  const timeLocale = localeFromEnvironment(facts.lc_all) || localeFromEnvironment(facts.lc_time) || locale
  const numberLocale = localeFromEnvironment(facts.lc_all) || localeFromEnvironment(facts.lc_numeric) || locale
  const name = (facts.full_name ?? "").trim() || nameFromAccount(facts.user ?? "")
  const separators = separatorsFor(numberLocale)
  return {
    ...DEFAULT_SETTINGS,
    name,
    initials: initialsOf(name),
    locale,
    dateFormat: datePatternFor(timeLocale),
    decimalSeparator: separators.decimal,
    groupSeparator: separators.group,
  }
}

/** The languages and regions offered in the setup and Options; anything else is typed in. (Inkspire's list.) */
export const LOCALE_OPTIONS = [
  "en-US",
  "en-GB",
  "en-AU",
  "en-CA",
  "en-IN",
  "fr-FR",
  "fr-CA",
  "de-DE",
  "de-CH",
  "es-ES",
  "es-MX",
  "it-IT",
  "pt-BR",
  "pt-PT",
  "nl-NL",
  "sv-SE",
  "da-DK",
  "nb-NO",
  "fi-FI",
  "pl-PL",
  "cs-CZ",
  "ru-RU",
  "tr-TR",
  "ar-SA",
  "he-IL",
  "hi-IN",
  "ja-JP",
  "ko-KR",
  "zh-CN",
  "zh-TW",
]

export const DATE_FORMAT_OPTIONS = [
  "M/d/yyyy",
  "MM/dd/yyyy",
  "d/M/yyyy",
  "dd/MM/yyyy",
  "dd.MM.yyyy",
  "d.M.yyyy",
  "yyyy-MM-dd",
  "yyyy/MM/dd",
  "dd-MM-yyyy",
]

/** A locale's name in English for the lists ("English (United Kingdom)"); the tag itself when unknown. */
export function localeLabel(tag: string): string {
  try {
    return new Intl.DisplayNames(["en"], { type: "language" }).of(tag) ?? tag
  } catch {
    return tag
  }
}
