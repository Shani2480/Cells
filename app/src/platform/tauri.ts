import type { AiKeyStatus, AiPort, Platform, ScreenshotPort, SystemFacts } from "@cells/platform"
import { invoke } from "@tauri-apps/api/core"
import { listen } from "@tauri-apps/api/event"
import { readText, writeHtml, writeText } from "@tauri-apps/plugin-clipboard-manager"
import { getCurrentWindow } from "@tauri-apps/api/window"
import { openPath, openUrl } from "@tauri-apps/plugin-opener"
import { onOpenedDocuments } from "./android-files"
import { tauriFiles } from "./tauri-files"
import { displayMediaScreenshot, webPlatform } from "./web"

/**
 * The window's listener for links and files. Requests are taken from the shell once; one that comes back after its
 * listener went away goes to the listener there is then (React's StrictMode starts effects twice in a dev build, so
 * the first listener takes the launch link and is gone before the answer comes), or waits for the next one.
 */
let listener: ((items: string[]) => void) | null = null
const unclaimed: string[] = []

function hand(items: string[]) {
  if (!items.length) return
  if (listener) listener(items)
  else unclaimed.push(...items)
}

export function isTauri() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window
}

/** The desktop and Android shells: the system clipboard through the clipboard plugin, real window buttons. */
export function tauriPlatform(): Platform {
  const appWindow = getCurrentWindow()
  return {
    os: webPlatform.os === "web" ? "linux" : webPlatform.os,
    files: tauriFiles(),
    opens: {
      onRequest(handler) {
        let disposed = false
        let unlisten: (() => void) | undefined
        listener = handler
        const take = () => void invoke<string[]>("take_pending_opens").then(hand)
        if (unclaimed.length) {
          const items = unclaimed.splice(0)
          queueMicrotask(() => hand(items))
        }
        void listen("cells-open", take).then((fn) => {
          unlisten = fn
          if (disposed) fn()
        })
        // What came with the launch, before the window listened.
        take()
        // Android: files from other apps ("Open with", Share) come through MainActivity.
        const stopAndroid = onOpenedDocuments(hand)
        return () => {
          disposed = true
          stopAndroid()
          if (listener === handler) listener = null
          unlisten?.()
        }
      },
      // Paths come resolved (app/src/sheet/hyperlinks.ts); anything with a scheme (https:, mailto:, content:) is a URL.
      async external(target) {
        if (/^[a-z][a-z0-9+.-]*:/i.test(target) && !/^[a-z]:[\\/]/i.test(target)) await openUrl(target)
        else await openPath(target)
      },
      async openCopy(name, bytes) {
        const path = await invoke<string>("write_object_copy", { name, bytes: Array.from(bytes) })
        await openPath(path)
      },
    },
    ai: tauriAi(),
    screenshot: tauriScreenshot(webPlatform.os),
    web: {
      async fetch(url) {
        const got = await invoke<{ bytes: number[]; contentType: string }>("fetch_url", { url })
        return { bytes: Uint8Array.from(got.bytes), type: got.contentType }
      },
    },
    links: {
      wanted: () => invoke<string | null>("read_link_wants"),
      publish: (name, text) => invoke<void>("write_link_snapshot", { name, text }),
      isEditCopy: (path) => invoke<boolean>("is_edit_copy", { path }).catch(() => false),
    },
    settings: {
      read: () => invoke<string | null>("read_shared_settings"),
      write: (text) => invoke<void>("write_shared_settings", { text }),
      location: () => invoke<string>("shared_settings_location").catch(() => ""),
      async systemFacts() {
        const intlLocale = typeof navigator === "undefined" ? undefined : navigator.language
        try {
          return { ...(await invoke<SystemFacts>("system_defaults")), intlLocale }
        } catch {
          return { intlLocale }
        }
      },
    },
    clipboard: {
      async write(content) {
        if (content.html) await writeHtml(content.html, content.text)
        else await writeText(content.text)
      },
      async read() {
        try {
          return { text: await readText() }
        } catch {
          return null
        }
      },
    },
    window: {
      controllable: webPlatform.os !== "android",
      minimize: () => appWindow.minimize(),
      toggleMaximize: () => appWindow.toggleMaximize(),
      close: () => appWindow.close(),
      isMaximized: () => appWindow.isMaximized(),
      setTitle: async (title) => {
        document.title = title
        await appWindow.setTitle(title)
      },
      onCloseRequested(allow) {
        let disposed = false
        let unlisten: (() => void) | undefined
        void appWindow
          .onCloseRequested(async (event) => {
            if (!(await allow())) event.preventDefault()
          })
          .then((fn) => {
            unlisten = fn
            if (disposed) fn()
          })
        return () => {
          disposed = true
          unlisten?.()
        }
      },
      onMaximizedChange(callback) {
        let disposed = false
        let unlisten: (() => void) | undefined
        void appWindow
          .onResized(() => {
            void appWindow.isMaximized().then((maximized) => {
              if (!disposed) callback(maximized)
            })
          })
          .then((fn) => {
            unlisten = fn
            if (disposed) fn()
          })
        return () => {
          disposed = true
          unlisten?.()
        }
      },
      onMenuCommand(callback) {
        let disposed = false
        let unlisten: (() => void) | undefined
        void listen<string>("menu-command", (event) => callback(event.payload)).then((fn) => {
          unlisten = fn
          if (disposed) fn()
        })
        return () => {
          disposed = true
          unlisten?.()
        }
      },
    },
  }
}

let requests = 0

/** AI requests through the shell (src-tauri/src/ai.rs): the reply streams back as `ai-chunk` events. */
function tauriAi(): AiPort {
  return {
    available: true,
    async generate(request, onText, signal) {
      const requestId = `cells-ai-${Date.now().toString(36)}-${++requests}`
      const unlisten = await listen<{ requestId: string; text: string }>("ai-chunk", (event) => {
        if (event.payload.requestId === requestId) onText(event.payload.text)
      })
      const stop = () => void invoke("ai_cancel", { requestId })
      signal?.addEventListener("abort", stop)
      try {
        if (signal?.aborted) throw new Error("Stopped.")
        return await invoke<string>("ai_generate", {
          requestId,
          system: request.system,
          prompt: request.prompt,
          context: request.context,
          providerConfig: request.target,
        })
      } catch (error) {
        throw new Error(typeof error === "string" ? error : error instanceof Error ? error.message : String(error), {
          cause: error,
        })
      } finally {
        signal?.removeEventListener("abort", stop)
        unlisten()
      }
    },
    test: (target) =>
      invoke<string>("ai_test", { providerConfig: target }).catch((error: unknown) => {
        throw new Error(String(error))
      }),
    keyStatus: () => invoke<AiKeyStatus>("ai_key_status").catch(() => null),
    saveKey: (provider, key) =>
      invoke<string>("ai_save_key", { provider, key }).catch((error: unknown) => {
        throw new Error(String(error))
      }),
  }
}

/**
 * Insert › Screenshot in the desktop app: the system's own tool on Linux (the desktop portal) and macOS
 * (screencapture), the webview's screen picker on Windows; none on Android.
 */
function tauriScreenshot(os: Platform["os"]): ScreenshotPort | undefined {
  if (os === "android") return undefined
  if (os === "windows") return displayMediaScreenshot
  return {
    async take() {
      const bytes = await invoke<number[] | null>("take_screenshot")
      return bytes ? imageDataUrl(Uint8Array.from(bytes)) : null
    },
  }
}

/** Image bytes as a data URL, its type from its first bytes (PNG unless it's a JPEG). */
export function imageDataUrl(bytes: Uint8Array): string {
  const type = bytes[0] === 0xff && bytes[1] === 0xd8 ? "image/jpeg" : "image/png"
  let binary = ""
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return `data:${type};base64,${btoa(binary)}`
}
