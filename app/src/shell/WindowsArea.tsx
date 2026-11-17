import { Maximize2, X } from "lucide-react"
import type { CSSProperties, ReactNode } from "react"
import { useSyncExternalStore } from "react"
import type { DocumentWindows } from "@/file/windows"
import type { SheetController } from "@/sheet/controller"

/**
 * The sheet area with View › New Window's windows. Maximized (Excel's default) it is just the active window's grid;
 * arranged (Arrange All, View Side by Side) each window is a panel with its title, Maximize and Close. Clicking a
 * panel makes it the active window, which the ribbon, formula bar, sheet tabs and status bar then follow.
 */
export function WindowsArea({
  windows,
  baseTitle,
  grid,
}: {
  windows: DocumentWindows
  /** The workbook's title ("Budget"); each panel shows "Budget:2". */
  baseTitle: string
  grid: (c: SheetController, active: boolean) => ReactNode
}) {
  useSyncExternalStore(windows.subscribe, windows.getVersion)
  const shown = windows.shown
  const active = windows.active
  if (!windows.arranged || shown.length < 2) return <>{grid(active, true)}</>

  const kind = windows.arrangement!
  const order = kind === "cascade" ? windows.cascadeOrder : shown
  const n = order.length
  const cols = kind === "vertical" ? n : kind === "horizontal" ? 1 : Math.ceil(Math.sqrt(n))
  const rows = Math.ceil(n / cols)
  const areaStyle: CSSProperties =
    kind === "cascade"
      ? {}
      : { gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))`, gridTemplateRows: `repeat(${rows}, minmax(0, 1fr))` }

  return (
    <div className={`win-area is-${kind}`} style={areaStyle}>
      {order.map((c, i) => {
        const isActive = c === active
        // Tiled: the last window takes the rest of its row.
        const lastSpan = kind === "tiled" && i === n - 1 ? cols * rows - n + 1 : 1
        const style: CSSProperties =
          kind === "cascade"
            ? ({ "--win-step": shown.indexOf(c), "--win-count": n, zIndex: i + 1 } as CSSProperties)
            : lastSpan > 1
              ? { gridColumn: `span ${lastSpan}` }
              : {}
        const title = windows.title(baseTitle, c)
        return (
          <section
            key={windows.key(c)}
            className={`win-panel${isActive ? " is-active" : ""}`}
            style={style}
            aria-label={title}
            onPointerDownCapture={() => windows.activate(c)}
          >
            <header className="win-title" onDoubleClick={() => windows.maximize()}>
              <span className="win-name">{title}</span>
              <span className="win-sheet">{c.sheet.name}</span>
              <button
                type="button"
                className="win-button"
                title="Maximize"
                aria-label={`Maximize ${title}`}
                onClick={() => {
                  windows.activate(c)
                  windows.maximize()
                }}
              >
                <Maximize2 size={12} strokeWidth={1.75} />
              </button>
              <button
                type="button"
                className="win-button is-close"
                title="Close Window"
                aria-label={`Close ${title}`}
                onClick={() => windows.close(c)}
              >
                <X size={14} strokeWidth={1.75} />
              </button>
            </header>
            <div className="win-body">{grid(c, isActive)}</div>
          </section>
        )
      })}
    </div>
  )
}
