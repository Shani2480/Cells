import type {
  AiPort,
  ClipboardContent,
  LinkPort,
  Platform,
  ScreenshotPort,
  SettingsPort,
  WebPort,
} from "@cells/platform"
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

/**
 * A screenshot through the webview's screen picker (getDisplayMedia): the person picks a screen, window or tab, and
 * one frame of it becomes the picture. Absent where the webview has no screen capture.
 */
export const displayMediaScreenshot: ScreenshotPort | undefined =
  typeof navigator !== "undefined" && typeof navigator.mediaDevices?.getDisplayMedia === "function"
    ? {
        async take() {
          let stream: MediaStream
          try {
            stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false })
          } catch (error) {
            // Cancelling the picker is a NotAllowedError; that's not a failure.
            if (error instanceof DOMException && error.name === "NotAllowedError") return null
            throw error
          }
          try {
            const video = document.createElement("video")
            video.muted = true
            video.srcObject = stream
            await video.play()
            // Let the first real frame arrive.
            await new Promise((r) => setTimeout(r, 150))
            const canvas = document.createElement("canvas")
            canvas.width = video.videoWidth
            canvas.height = video.videoHeight
            canvas.getContext("2d")!.drawImage(video, 0, 0)
            return canvas.toDataURL("image/png")
          } finally {
            for (const track of stream.getTracks()) track.stop()
          }
        },
      }
    : undefined

/** The browser's fetch: only sites that allow other pages to read them (CORS) work here. */
export const browserWeb: WebPort = {
  async fetch(url) {
    let response: Response
    try {
      response = await fetch(url)
    } catch {
      throw new Error(
        "Couldn't read the address. In a browser, only sites that allow it can be read; the Cells app reads any.",
      )
    }
    if (!response.ok) throw new Error(`The site answered ${response.status} ${response.statusText}.`)
    return { bytes: new Uint8Array(await response.arrayBuffer()), type: response.headers.get("content-type") ?? "" }
  },
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
  opens: { onRequest: () => () => {}, external: openInBrowser, openCopy: downloadCopy },
  links: webLinks(),
  ai: noAi,
  screenshot: displayMediaScreenshot,
  web: browserWeb,
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

/** A browser tab opens web and mail addresses in a new tab; files on the disk are out of its reach. */
/** A browser can't hand a file to another app: the embedded file downloads, for the person to open. */
export async function downloadCopy(name: string, bytes: Uint8Array): Promise<void> {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart]))
  const a = document.createElement("a")
  a.href = url
  a.download = name || "Object"
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 60_000)
}

export async function openInBrowser(target: string): Promise<void> {
  if (!/^(https?|ftp|mailto):/i.test(target)) throw new Error("A browser can't open files on this computer.")
  const opened = window.open(target, "_blank", "noopener,noreferrer")
  // With noopener the call returns null even when it opened; a blocked pop-up can't be told apart.
  void opened
}
