/**
 * The person's settings as the app uses them: loaded from the file shared with Inkspire, saved back into it, and
 * read again when the window comes back to the front (Inkspire may have changed them meanwhile).
 */
import type { SettingsPort } from "@cells/platform"
import { create } from "zustand"
import { DEFAULT_SETTINGS, mergeShared, parseShared, type SharedSettings } from "./shared"

type SettingsState = SharedSettings & {
  /** Loaded from the file at least once (the setup waits for it, so it doesn't flash before a finished setup). */
  loaded: boolean
}

export const useSettings = create<SettingsState>(() => ({ ...DEFAULT_SETTINGS, loaded: false }))

let port: SettingsPort | null = null
/** Saves run one after another, so two quick changes can't overwrite each other's read of the file. */
let queue: Promise<void> = Promise.resolve()

const pick = (s: SharedSettings): SharedSettings => ({
  name: s.name,
  initials: s.initials,
  locale: s.locale,
  dateFormat: s.dateFormat,
  decimalSeparator: s.decimalSeparator,
  groupSeparator: s.groupSeparator,
  theme: s.theme,
  setupDone: s.setupDone,
  aiEnabled: s.aiEnabled,
  aiProvider: s.aiProvider,
  aiModels: s.aiModels,
  aiAddresses: s.aiAddresses,
  dictionaryWords: s.dictionaryWords,
})

/** Reads the shared file (on start, and when the window is focused again). */
export async function loadSettings(from: SettingsPort): Promise<void> {
  port = from
  let source: string | null = null
  try {
    source = await from.read()
  } catch {
    // An unreadable file reads as no settings; the next save writes a fresh one.
  }
  const next = parseShared(source)
  const now = useSettings.getState()
  if (now.loaded && JSON.stringify(pick(now)) === JSON.stringify(next)) return
  useSettings.setState({ ...next, loaded: true })
}

/**
 * Changes some settings and saves just those into the shared file, over what it holds now; what the other app
 * changed meanwhile is taken in.
 */
export function updateSettings(patch: Partial<SharedSettings>): Promise<void> {
  useSettings.setState(patch)
  const target = port
  if (!target) return Promise.resolve()
  queue = queue.then(async () => {
    let source: string | null
    try {
      source = await target.read()
    } catch {
      source = null
    }
    const merged = mergeShared(source, patch)
    try {
      await target.write(merged)
    } catch {
      // Kept for this session; the next change tries to save again.
      return
    }
    useSettings.setState({ ...parseShared(merged), loaded: true })
  })
  return queue
}

/** The name notes and comments are signed with: the person's, or Excel's fallback "User" until one is set. */
export function authorName(): string {
  return useSettings.getState().name.trim() || "User"
}
