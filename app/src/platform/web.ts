import type { AiPort, ClipboardContent, LinkPort, Platform, SettingsPort } from "@cells/platform"
import { webFiles } from "./web-files"

function detectOs(): Platform["os"] {
  if (typeof navigator === "undefined") return "web"
  const agent = navigator.userAgent
  if (/Android/i.test(agent)) return "android"
  if (/Mac|iPhone|iPad/i.test(agent)) return "macos"
  if (/Windows/i.test(agent)) return "windows"
  if (/Linux/i.test(agent)) return "linux"
  return "web"
}

/** The browser build (dev server and tests): the async clipboard where allowed, no window controls. */
/** A browser can't keep keys safe or reach local models: the AI tools need the desktop app. */
export const noAi: AiPort = {
  available: false,
  generate: async () => {
    throw new Error("The AI tools need the Cells desktop app.")
  },
  test: async () => {
    throw new Error("The AI tools need the Cells desktop app.")
  },
  keyStatus: async () => null,
  saveKey: async () => {
    throw new Error("The AI tools need the Cells desktop app.")
  },
}
export const webPlatform: Platform = {
  os: detectOs(),
  clipboard: {
    async write(content: ClipboardContent) {
      if (typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
        const items: Record<string, Blob> = { "text/plain": new Blob([content.text], { type: "text/plain" }) }
        if (content.html) items["text/html"] = new Blob([content.html], { type: "text/html" })
        await navigator.clipboard.write([new ClipboardItem(items)])
        return
      }
      await navigator.clipboard?.writeText(content.text)
    },
    async read() {
      try {
        const text = await navigator.clipboard.readText()
        return { text }
      } catch {
        return null
      }
    },
  },
  window: {
    controllable: false,
    async minimize() {},
    async toggleMaximize() {},
    async close() {},
    async isMaximized() {
      return false
    },
    async setTitle(title: string) {
      document.title = title
    },
    onMaximizedChange() {
      return () => {}
    },
    onCloseRequested(allow) {
      const listener = (event: BeforeUnloadEvent) => {
        // The browser's own "Leave site?" prompt; only a synchronous answer counts here.
        const answer = allow()
        if (answer === false) event.preventDefault()
      }
      window.addEventListener("beforeunload", listener)
      return () => window.removeEventListener("beforeunload", listener)
    },
  },
  files: webFiles(),
  settings: webSettings(),
  // A browser tab is never handed links or files by the system.
  opens: { onRequest: () => () => {} },
  links: webLinks(),
  ai: noAi,
}

/** The browser build: the same files as storage entries (nothing outside the tab reads them). */
export function webLinks(): LinkPort {
  const storage = () => {
    try {
      return window.localStorage
    } catch {
      return null
    }
  }
  return {
    wanted: async () => storage()?.getItem("inkspire.links.wanted") ?? null,
    async publish(name, text) {
      try {
        storage()?.setItem(`inkspire.links.${name}`, text)
      } catch {
        // Full or blocked storage: nothing to tell.
      }
    },
    isEditCopy: async () => false,
  }
}

const SHARED_KEY = "inkspire.shared-settings"

/** The browser build has no file both apps can see: the shared settings live in this origin's storage. */
export function webSettings(): SettingsPort {
  return {
    async read() {
      try {
        return localStorage.getItem(SHARED_KEY)
      } catch {
        return null
      }
    },
    async write(text) {
      try {
        localStorage.setItem(SHARED_KEY, text)
      } catch {
        // Without storage the settings last as long as the page.
      }
    },
    async location() {
      return ""
    },
    async systemFacts() {
      return { intlLocale: typeof navigator === "undefined" ? undefined : navigator.language }
    },
  }
}
