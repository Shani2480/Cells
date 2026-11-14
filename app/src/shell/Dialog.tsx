import { X } from "lucide-react"
import { useEffect, useId, useRef, type ReactNode } from "react"
import { createPortal } from "react-dom"

/**
 * The one dialog frame: title bar in the window's title colour with Close, modal, Esc cancels, Enter accepts,
 * focus starts on the first field. Mounted on document.body.
 */
export function Dialog({
  title,
  onCancel,
  onAccept,
  children,
  width = 320,
  acceptLabel = "OK",
}: {
  title: string
  onCancel: () => void
  onAccept: () => void
  children: ReactNode
  width?: number
  acceptLabel?: string
}) {
  const titleId = useId()
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const first = ref.current?.querySelector<HTMLElement>("input, select, button.dialog-list-item, button.dialog-btn")
    first?.focus()
    if (first instanceof HTMLInputElement) first.select()
  }, [])

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault()
        event.stopPropagation()
        onCancel()
      }
    }
    window.addEventListener("keydown", onKey, true)
    return () => window.removeEventListener("keydown", onKey, true)
  }, [onCancel])

  return createPortal(
    <div
      className="dialog-backdrop"
      onPointerDown={(event) => event.target === event.currentTarget && event.preventDefault()}
    >
      <div
        ref={ref}
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        style={{ width }}
        onKeyDown={(event) => {
          if (event.key === "Enter" && !(event.target instanceof HTMLButtonElement && event.target.type === "button")) {
            event.preventDefault()
            onAccept()
          }
        }}
      >
        <div className="dialog-title">
          <span id={titleId}>{title}</span>
          <button type="button" className="dialog-title-close" aria-label="Close" title="Close" onClick={onCancel}>
            <X size={14} strokeWidth={1.75} />
          </button>
        </div>
        <div className="dialog-body">{children}</div>
        <div className="dialog-actions">
          <button type="button" className="dialog-btn is-primary" onClick={onAccept}>
            {acceptLabel}
          </button>
          <button type="button" className="dialog-btn" onClick={onCancel}>
            Cancel
          </button>
        </div>
      </div>
    </div>,
    document.body,
  )
}
