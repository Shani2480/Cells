import type { Platform, SystemFacts } from "@cells/platform"
import { invoke } from "@tauri-apps/api/core"
import { readText, writeHtml, writeText } from "@tauri-apps/plugin-clipboard-manager"
import { getCurrentWindow } from "@tauri-apps/api/window"
import { tauriFiles } from "./tauri-files"
import { webPlatform } from "./web"

export function isTauri() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window
}

/** The desktop and Android shells: the system clipboard through the clipboard plugin, real window buttons. */
export function tauriPlatform(): Platform {
  const appWindow = getCurrentWindow()
  return {
    os: webPlatform.os === "web" ? "linux" : webPlatform.os,
    files: tauriFiles(),
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
    },
  }
}
