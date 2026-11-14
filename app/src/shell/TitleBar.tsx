import { Minus, Redo2, Square, SquareSquare, Undo2, X } from "lucide-react"
import { useEffect, useState, useSyncExternalStore } from "react"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"

/**
 * Excel's title bar: Quick Access (Undo, Redo) at the left, the workbook's name in the middle, window buttons at
 * the right. The whole bar drags the window (Tauri's drag script, "deep" so the gaps count too).
 */
export function TitleBar({ controller, title }: { controller: SheetController; title: string }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const platform = usePlatform()
  const [maximized, setMaximized] = useState(false)
  const shortcut = (key: string) => (platform.os === "macos" ? `⌘${key}` : `Ctrl+${key}`)

  useEffect(() => {
    void platform.window.setTitle(title).catch(() => undefined)
  }, [platform, title])

  useEffect(() => {
    if (!platform.window.controllable) return
    void platform.window.isMaximized().then((value) => {
      setMaximized(value)
      document.documentElement.classList.toggle("is-maximized", value)
    })
    return platform.window.onMaximizedChange((value) => {
      setMaximized(value)
      document.documentElement.classList.toggle("is-maximized", value)
    })
  }, [platform])

  const undoLabel = controller.wb.undoLabel
  const redoLabel = controller.wb.redoLabel
  return (
    <header className="titlebar" data-tauri-drag-region="deep">
      <div className="titlebar-start">
        <img className="titlebar-logo" src="/logo-mark.svg" alt="" draggable={false} />
        <button
          type="button"
          className="titlebar-button"
          aria-label="Undo"
          title={undoLabel ? `Undo ${undoLabel} (${shortcut("Z")})` : `Undo (${shortcut("Z")})`}
          disabled={!controller.wb.canUndo}
          onClick={() => controller.undo()}
        >
          <Undo2 size={16} strokeWidth={1.75} />
        </button>
        <button
          type="button"
          className="titlebar-button"
          aria-label="Redo"
          title={redoLabel ? `Redo ${redoLabel} (${shortcut("Y")})` : `Redo (${shortcut("Y")})`}
          disabled={!controller.wb.canRedo}
          onClick={() => controller.redo()}
        >
          <Redo2 size={16} strokeWidth={1.75} />
        </button>
      </div>
      <div className="titlebar-title">{title}</div>
      <div className="titlebar-end">
        {platform.window.controllable ? (
          <>
            <button
              type="button"
              className="window-button"
              aria-label="Minimize"
              title="Minimize"
              onClick={() => void platform.window.minimize()}
            >
              <Minus size={15} strokeWidth={1.5} />
            </button>
            <button
              type="button"
              className="window-button"
              aria-label={maximized ? "Restore" : "Maximize"}
              title={maximized ? "Restore" : "Maximize"}
              onClick={() => void platform.window.toggleMaximize()}
            >
              {maximized ? <SquareSquare size={13} strokeWidth={1.5} /> : <Square size={12} strokeWidth={1.5} />}
            </button>
            <button
              type="button"
              className="window-button is-close"
              aria-label="Close"
              title="Close"
              onClick={() => void platform.window.close()}
            >
              <X size={16} strokeWidth={1.5} />
            </button>
          </>
        ) : null}
      </div>
    </header>
  )
}
