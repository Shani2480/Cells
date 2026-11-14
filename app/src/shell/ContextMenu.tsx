import { ClipboardPaste, Copy, Scissors } from "lucide-react"
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"

export type MenuItem =
  | { kind: "separator" }
  | {
      kind: "item"
      label: string
      icon?: ReactNode
      shortcut?: string
      disabled?: boolean
      onSelect: () => void
    }

export const menuIcons = {
  cut: <Scissors size={15} strokeWidth={1.6} />,
  copy: <Copy size={15} strokeWidth={1.6} />,
  paste: <ClipboardPaste size={15} strokeWidth={1.6} />,
}

/** A right-click menu at a point, kept inside the window. Closes at once on choice, Esc or a press elsewhere. */
export function ContextMenu({
  x,
  y,
  items,
  onClose,
}: {
  x: number
  y: number
  items: MenuItem[]
  onClose: () => void
}) {
  const ref = useRef<HTMLDivElement>(null)
  const [place, setPlace] = useState({ left: x, top: y, visible: false })

  useLayoutEffect(() => {
    const menu = ref.current
    if (!menu) return
    const rect = menu.getBoundingClientRect()
    const left = Math.min(x, window.innerWidth - rect.width - 4)
    const top = y + rect.height > window.innerHeight - 4 ? Math.max(4, y - rect.height) : y
    setPlace({ left: Math.max(4, left), top, visible: true })
    menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus({ preventScroll: true })
  }, [x, y])

  useEffect(() => {
    const onDown = (event: PointerEvent) => {
      if (!ref.current?.contains(event.target as Node)) onClose()
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault()
        event.stopPropagation()
        onClose()
      }
    }
    window.addEventListener("pointerdown", onDown, true)
    window.addEventListener("keydown", onKey, true)
    window.addEventListener("blur", onClose)
    return () => {
      window.removeEventListener("pointerdown", onDown, true)
      window.removeEventListener("keydown", onKey, true)
      window.removeEventListener("blur", onClose)
    }
  }, [onClose])

  const move = (event: React.KeyboardEvent, delta: number) => {
    const buttons = [...(ref.current?.querySelectorAll<HTMLButtonElement>("button:not(:disabled)") ?? [])]
    const index = buttons.indexOf(document.activeElement as HTMLButtonElement)
    buttons[(index + delta + buttons.length) % buttons.length]?.focus()
    event.preventDefault()
  }

  return createPortal(
    <div
      ref={ref}
      className="context-menu"
      role="menu"
      style={{ left: place.left, top: place.top, visibility: place.visible ? "visible" : "hidden" }}
      onKeyDown={(event) => {
        if (event.key === "ArrowDown") move(event, 1)
        if (event.key === "ArrowUp") move(event, -1)
      }}
      onContextMenu={(event) => event.preventDefault()}
    >
      {items.map((item, index) =>
        item.kind === "separator" ? (
          <div key={index} className="menu-separator" role="separator" />
        ) : (
          <button
            key={index}
            type="button"
            role="menuitem"
            className="menu-row"
            disabled={item.disabled}
            onClick={() => {
              onClose()
              item.onSelect()
            }}
          >
            <span className="menu-row-icon">{item.icon}</span>
            <span>{item.label}</span>
            {item.shortcut ? <span className="menu-row-shortcut">{item.shortcut}</span> : null}
          </button>
        ),
      )}
    </div>,
    document.body,
  )
}
