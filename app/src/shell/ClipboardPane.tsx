import { BarChart3, ClipboardPaste, Trash2, X } from "lucide-react"
import { useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"

/** A copied item's preview: its first rows, cells apart by spaces, cut short. */
export function clipPreview(text: string): string {
  const lines = text
    .replace(/\r/g, "")
    .split("\n")
    .filter((line) => line.replace(/\t/g, "").trim())
    .slice(0, 3)
  const shown = lines.map((l) => l.split("\t").join("  ")).join("\n")
  return shown.length > 160 ? shown.slice(0, 159) + "…" : shown
}

/**
 * Home › Clipboard (the launcher): the Office Clipboard, the last 24 things copied in Cells, newest first. Clicking
 * one pastes it at the selection; Paste All pastes them all, oldest first, one under another; Clear All empties it.
 * Copies made in other apps aren't collected (Cells doesn't watch the system clipboard).
 */
export function ClipboardPane({
  controller,
  onPaste,
  onDone,
}: {
  controller: SheetController
  /** Pastes text and HTML at the selection, as Ctrl+V does. */
  onPaste: (text: string, html: string) => void
  onDone: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const items = controller.clipboardHistory
  const close = () => {
    controller.toggleClipboardPane(false)
    onDone()
  }
  return (
    <aside
      className="comments-pane clipboard-pane"
      aria-label="Clipboard"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Clipboard</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.6} aria-hidden />
        </button>
      </header>
      <div className="clipboard-pane-actions">
        <button
          type="button"
          className="comments-pane-new"
          disabled={!items.some((i) => !i.chart)}
          onClick={() => {
            const texts = [...items].reverse().filter((i) => !i.chart)
            onPaste(texts.map((i) => i.text.replace(/\r?\n$/, "")).join("\n"), "")
            onDone()
          }}
        >
          <ClipboardPaste size={16} strokeWidth={1.6} aria-hidden />
          Paste All
        </button>
        <button
          type="button"
          className="comments-pane-new"
          disabled={!items.length}
          onClick={() => controller.forgetCopied()}
        >
          <Trash2 size={16} strokeWidth={1.6} aria-hidden />
          Clear All
        </button>
      </div>
      <div className="comments-pane-list">
        <p className="comments-pane-empty">
          {items.length ? "Click an item to paste:" : "Copy something and it shows here, ready to paste again."}
        </p>
        {items.map((item, index) => (
          <div key={index} className="clipboard-item">
            <button
              type="button"
              className="comments-pane-thread clipboard-item-paste"
              aria-label={`Paste ${item.chart ? "chart" : clipPreview(item.text)}`}
              onClick={() => {
                onPaste(item.text, item.html)
                onDone()
              }}
            >
              {item.chart ? (
                <span className="clipboard-item-chart">
                  <BarChart3 size={16} strokeWidth={1.6} aria-hidden /> Chart
                </span>
              ) : (
                <span className="comments-pane-text">{clipPreview(item.text)}</span>
              )}
            </button>
            <button
              type="button"
              className="comments-pane-close clipboard-item-delete"
              aria-label="Delete"
              title="Delete"
              onClick={() => controller.forgetCopied(index)}
            >
              <X size={14} strokeWidth={1.6} aria-hidden />
            </button>
          </div>
        ))}
      </div>
    </aside>
  )
}
