import { Minus, Square, SquareSquare, X } from "lucide-react"
import { useEffect, useState, useSyncExternalStore, type ReactNode } from "react"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"

/**
 * Excel's title bar: the Quick Access Toolbar and the workbook's name at the left, Search in the middle, window
 * buttons at the right. The whole bar drags the window (Tauri's drag script, "deep" so the gaps count too).
 */
export function TitleBar({
  controller,
  title,
  qat,
  search,
}: {
  controller: SheetController
  title: string
  /** The Quick Access Toolbar. */
  qat?: ReactNode
  /** Search (Tell Me). */
  search?: ReactNode
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const platform = usePlatform()
  const [maximized, setMaximized] = useState(false)

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

  return (
    <header className="titlebar" data-tauri-drag-region="deep">
      <div className="titlebar-start">
        <img className="titlebar-logo" src="/logo-mark.svg" alt="" draggable={false} />
        {qat}
        <div className="titlebar-title">{title}</div>
      </div>
      <div className="titlebar-center">{search}</div>
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
