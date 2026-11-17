import { ChevronRight, ClipboardPaste, Copy, Link, Scissors, SquarePen } from "lucide-react"
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
      /** Tooltip (a disabled item says why: "Not built yet."). */
      title?: string
      /** Opens more choices (drawn with a ›). */
      more?: boolean
      /** A submenu (Filter ›, Sort ›): opens beside the row on hover, a click or Right. */
      items?: MenuItem[]
      onSelect: () => void
    }

export const menuIcons = {
  cut: <Scissors size={15} strokeWidth={1.6} />,
  copy: <Copy size={15} strokeWidth={1.6} />,
  paste: <ClipboardPaste size={15} strokeWidth={1.6} />,
  formatCells: <SquarePen size={15} strokeWidth={1.6} />,
  link: <Link size={15} strokeWidth={1.6} />,
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

  // The open submenu: its row's index and where it goes.
  const [sub, setSub] = useState<{ index: number; left: number; top: number } | null>(null)
  const subRef = useRef<HTMLDivElement>(null)
  const openSub = (index: number, row: HTMLElement, focus: boolean) => {
    const rect = row.getBoundingClientRect()
    setSub({ index, left: rect.right - 2, top: rect.top - 4 })
    if (focus)
      requestAnimationFrame(() => subRef.current?.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus())
  }
  // Kept inside the window: to the left of the menu when there's no room on the right.
  useLayoutEffect(() => {
    const el = subRef.current
    if (!el || !sub) return
    const rect = el.getBoundingClientRect()
    const menu = ref.current!.getBoundingClientRect()
    if (rect.right > window.innerWidth - 4) el.style.left = `${Math.max(4, menu.left - rect.width + 2)}px`
    if (rect.bottom > window.innerHeight - 4) el.style.top = `${Math.max(4, window.innerHeight - 4 - rect.height)}px`
  }, [sub])

  const rows = (root: HTMLElement | null) => [
    ...(root?.querySelectorAll<HTMLButtonElement>(":scope > .menu-row-wrap > button:not(:disabled)") ?? []),
  ]
  const move = (event: React.KeyboardEvent, delta: number) => {
    const inSub = subRef.current?.contains(document.activeElement)
    const buttons = rows(inSub ? subRef.current : ref.current)
    const index = buttons.indexOf(document.activeElement as HTMLButtonElement)
    buttons[(index + delta + buttons.length) % buttons.length]?.focus()
    event.preventDefault()
  }

  const list = (entries: MenuItem[], nested: boolean) =>
    entries.map((item, index) =>
      item.kind === "separator" ? (
        <div key={index} className="menu-separator" role="separator" />
      ) : (
        <span key={index} className="menu-row-wrap" title={item.title}>
          <button
            type="button"
            role="menuitem"
            className="menu-row"
            disabled={item.disabled}
            aria-haspopup={item.more || item.items ? "menu" : undefined}
            aria-expanded={item.items ? sub?.index === index : undefined}
            onPointerEnter={(event) => {
              if (nested) return
              if (item.items && !item.disabled) openSub(index, event.currentTarget, false)
              else setSub(null)
            }}
            onKeyDown={(event) => {
              if (!nested && item.items && event.key === "ArrowRight") {
                event.preventDefault()
                openSub(index, event.currentTarget, true)
              }
            }}
            onClick={(event) => {
              if (item.items) {
                if (!nested) openSub(index, event.currentTarget, true)
                return
              }
              onClose()
              item.onSelect()
            }}
          >
            <span className="menu-row-icon">{item.icon}</span>
            <span>{item.label}</span>
            {item.shortcut ? <span className="menu-row-shortcut">{item.shortcut}</span> : null}
            {item.more || item.items ? (
              <span className="menu-row-shortcut">
                <ChevronRight size={14} strokeWidth={1.6} />
              </span>
            ) : null}
          </button>
        </span>
      ),
    )

  const open = sub ? items[sub.index] : null
  return createPortal(
    <div
      ref={ref}
      className="context-menu"
      role="menu"
      style={{ left: place.left, top: place.top, visibility: place.visible ? "visible" : "hidden" }}
      onKeyDown={(event) => {
        if (event.key === "ArrowDown") move(event, 1)
        if (event.key === "ArrowUp") move(event, -1)
        if (event.key === "ArrowLeft" && subRef.current?.contains(document.activeElement)) {
          event.preventDefault()
          rows(ref.current)
            .find((b) => b.getAttribute("aria-expanded") === "true")
            ?.focus()
          setSub(null)
        }
      }}
      onContextMenu={(event) => event.preventDefault()}
    >
      {list(items, false)}
      {open?.kind === "item" && open.items ? (
        <div
          ref={subRef}
          className="context-menu is-submenu"
          role="menu"
          aria-label={open.label}
          style={{ position: "fixed", left: sub!.left, top: sub!.top }}
        >
          {list(open.items, true)}
        </div>
      ) : null}
    </div>,
    document.body,
  )
}
