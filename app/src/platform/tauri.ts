import type { Platform } from "@cells/platform"
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
