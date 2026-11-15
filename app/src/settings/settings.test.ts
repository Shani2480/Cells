import { beforeEach, describe, expect, it } from "vitest"
import type { SettingsPort } from "@cells/platform"
import { authorName, loadSettings, updateSettings, useSettings } from "./settings"
import { DEFAULT_SETTINGS } from "./shared"

/** A settings file both apps write to, held in memory. */
function sharedFile(initial: string | null = null) {
  const file = { text: initial }
  const port: SettingsPort = {
    read: async () => file.text,
    write: async (text) => {
      file.text = text
    },
    location: async () => "~/.config/inkspire/settings.json",
    systemFacts: async () => ({}),
  }
  return { file, port }
}

beforeEach(() => useSettings.setState({ ...DEFAULT_SETTINGS, loaded: false }))

describe("settings store", () => {
  it("loads the shared file and signs notes with the name in it", async () => {
    const { port } = sharedFile(JSON.stringify({ user: { name: "Alex Demo" } }))
    expect(authorName()).toBe("User")
    await loadSettings(port)
    expect(useSettings.getState()).toMatchObject({ loaded: true, name: "Alex Demo" })
    expect(authorName()).toBe("Alex Demo")
  })

  it("saves into what the file holds now, so a change Inkspire made meanwhile survives", async () => {
    const { file, port } = sharedFile(JSON.stringify({ user: { name: "Alex Demo" } }))
    await loadSettings(port)
    // Inkspire changes the theme and adds a key of its own while Cells is open.
    file.text = JSON.stringify({ user: { name: "Alex Demo" }, appearance: { theme: "dark" }, inkspireOnly: 1 })
    await updateSettings({ initials: "AXD" })
    const saved = JSON.parse(file.text!)
    expect(saved.user).toEqual({ name: "Alex Demo", initials: "AXD" })
    expect(saved.inkspireOnly).toBe(1)
    // Only the initials were written: Inkspire's new theme stays, and Cells takes it in.
    expect(saved.appearance.theme).toBe("dark")
    expect(useSettings.getState().theme).toBe("dark")
  })

  it("picks up a change from the other app when loaded again (window focus)", async () => {
    const { file, port } = sharedFile(JSON.stringify({ appearance: { theme: "light" } }))
    await loadSettings(port)
    file.text = JSON.stringify({ appearance: { theme: "dark" } })
    await loadSettings(port)
    expect(useSettings.getState().theme).toBe("dark")
  })
})
