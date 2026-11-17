import { ChevronLeft, ChevronRight, Plus } from "lucide-react"
import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react"
import { ColorPanel } from "@/ribbon/colors"
import type { SheetController } from "@/sheet/controller"
import { ContextMenu, type MenuItem } from "./ContextMenu"

const NOT_BUILT = "Not built yet."
/** Pixels a press must travel before it drags a tab (a shorter move is still a click). */
const DRAG_START = 4

type Drag = { index: number; startX: number; moved: boolean; before: number | null; copy: boolean }

/**
 * Excel's sheet tab bar: scroll arrows, one tab per visible sheet (colour, active underline), the New Sheet button.
 * Click to show a sheet, double-click to rename, drag to move (Ctrl+drag copies), right-click for the tab menu.
 */
export function SheetTabs({
  controller,
  onDialog,
  onFocusGrid,
}: {
  controller: SheetController
  /** Opens the shell's dialogs: "moveCopySheet", "unhideSheet". */
  onDialog: (name: string) => void
  /** Keys go back to the sheet after a click on a tab. */
  onFocusGrid: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const sheets = controller.wb.sheets
  const commands = controller.sheets
  const strip = useRef<HTMLDivElement>(null)
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null)
  const [colorAt, setColorAt] = useState<{ x: number; y: number; index: number } | null>(null)
  const [drag, setDrag] = useState<Drag | null>(null)
  const dragRef = useRef<Drag | null>(null)
  const active = controller.sheetIndex

  // The active tab stays in view (Ctrl+PgDn, undo, a new sheet).
  useLayoutEffect(() => {
    const el = strip.current?.querySelector<HTMLElement>(`[data-sheet="${active}"]`)
    const box = strip.current
    if (!el || !box) return
    if (el.offsetLeft < box.scrollLeft) box.scrollLeft = el.offsetLeft
    else if (el.offsetLeft + el.offsetWidth > box.scrollLeft + box.clientWidth)
      box.scrollLeft = el.offsetLeft + el.offsetWidth - box.clientWidth
  }, [active, sheets.length])

  const scrollTabs = (dir: 1 | -1, toEnd: boolean) => {
    const box = strip.current
    if (!box) return
    if (toEnd) {
      box.scrollLeft = dir < 0 ? 0 : box.scrollWidth
      return
    }
    // One tab at a time: to the next tab edge past the current scroll position.
    const tabs = [...box.querySelectorAll<HTMLElement>("[data-sheet]")]
    if (dir > 0) {
      const next = tabs.find((t) => t.offsetLeft > box.scrollLeft + 1)
      if (next) box.scrollLeft = next.offsetLeft
    } else {
      const prev = [...tabs].reverse().find((t) => t.offsetLeft < box.scrollLeft - 1)
      box.scrollLeft = prev ? prev.offsetLeft : 0
    }
  }

  const tabMenu = (index: number, x: number, y: number) => {
    // A right-click inside the group keeps it (the menu then acts on every grouped sheet).
    if (!controller.group.has(index)) controller.activate(index)
    const grouped = controller.group.active
    // Review › Protect Workbook greys out what changes the structure (Excel).
    const locked = commands.isWorkbookProtected()
    const items: MenuItem[] = [
      { kind: "item", label: "Insert...", disabled: locked, onSelect: () => commands.insert(index) },
      { kind: "item", label: "Delete", disabled: locked, onSelect: () => commands.remove(index) },
      { kind: "item", label: "Rename", disabled: locked || grouped, onSelect: () => commands.beginRename(index) },
      {
        kind: "item",
        label: "Move or Copy...",
        disabled: locked || grouped,
        onSelect: () => onDialog("moveCopySheet"),
      },
      { kind: "item", label: "View Code", disabled: true, title: NOT_BUILT, onSelect: () => undefined },
      {
        kind: "item",
        label: commands.isProtected(index) ? "Unprotect Sheet..." : "Protect Sheet...",
        disabled: grouped,
        onSelect: () => {
          controller.activate(index)
          if (!commands.isProtected(index)) onDialog("protectSheet")
          else if (commands.needsPassword(index)) onDialog("unprotectSheet")
          else commands.unprotect()
        },
      },
      { kind: "item", label: "Tab Color", more: true, onSelect: () => setColorAt({ x, y, index }) },
      { kind: "separator" },
      { kind: "item", label: "Hide", disabled: locked, onSelect: () => commands.hide(index) },
      {
        kind: "item",
        label: "Unhide...",
        disabled: locked || commands.hidden().length === 0,
        onSelect: () => onDialog("unhideSheet"),
      },
      { kind: "separator" },
      grouped
        ? { kind: "item", label: "Ungroup Sheets", onSelect: () => controller.group.clear() }
        : { kind: "item", label: "Select All Sheets", onSelect: () => controller.group.selectAll() },
    ]
    setMenu({ x, y, items })
  }

  /** Right-click on the scroll arrows: every visible sheet, to jump to (Excel's Activate list). */
  const activateMenu = (x: number, y: number) => {
    const items: MenuItem[] = commands.visible().map((i) => ({
      kind: "item",
      label: sheets[i].name,
      icon: i === active ? "✓" : undefined,
      onSelect: () => controller.activate(i),
    }))
    setMenu({ x, y, items })
  }

  /** Where a drop at clientX lands: before the tab whose middle is right of it (sheets.length: at the end). */
  const dropTarget = (clientX: number): number => {
    const tabs = [...(strip.current?.querySelectorAll<HTMLElement>("[data-sheet]") ?? [])]
    for (const tab of tabs) {
      const r = tab.getBoundingClientRect()
      if (clientX < r.left + r.width / 2) return Number(tab.dataset.sheet)
    }
    return sheets.length
  }

  const onPointerDown = (event: React.PointerEvent, index: number) => {
    if (event.button !== 0 || commands.renaming === index) return
    const ctrl = event.ctrlKey || event.metaKey
    // Shift+click and Ctrl+click group sheets (no drag then).
    if (event.shiftKey || ctrl) {
      controller.group.click(index, { ctrl, shift: event.shiftKey })
      return
    }
    controller.group.click(index, { ctrl: false, shift: false })
    const d: Drag = { index, startX: event.clientX, moved: false, before: null, copy: false }
    dragRef.current = d
    ;(event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId)
  }

  const onPointerMove = (event: React.PointerEvent) => {
    const d = dragRef.current
    if (!d) return
    if (!d.moved && Math.abs(event.clientX - d.startX) < DRAG_START) return
    const next: Drag = { ...d, moved: true, before: dropTarget(event.clientX), copy: event.ctrlKey || event.metaKey }
    dragRef.current = next
    setDrag(next)
  }

  const onPointerUp = () => {
    const d = dragRef.current
    dragRef.current = null
    setDrag(null)
    if (commands.renaming === null) onFocusGrid()
    if (!d?.moved || d.before === null) return
    if (d.copy) commands.copy(d.index, d.before)
    else commands.move(d.index, d.before)
  }

  // Esc ends a drag without moving anything.
  useEffect(() => {
    if (!drag) return
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        dragRef.current = null
        setDrag(null)
      }
    }
    window.addEventListener("keydown", onKey, true)
    return () => window.removeEventListener("keydown", onKey, true)
  }, [drag])

  const marker =
    drag?.moved && drag.before !== null && (drag.copy || (drag.before !== drag.index && drag.before !== drag.index + 1))

  return (
    <div className="sheet-tabs" role="toolbar" aria-label="Sheet tabs">
      <div className="sheet-nav" onContextMenu={(e) => (e.preventDefault(), activateMenu(e.clientX, e.clientY))}>
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll to the previous sheet"
          title="Scroll to the previous sheet. Ctrl+Click to scroll to the first sheet. Right-click to see all sheets."
          onClick={(e) => scrollTabs(-1, e.ctrlKey || e.metaKey)}
        >
          <ChevronLeft size={14} strokeWidth={1.75} />
        </button>
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll to the next sheet"
          title="Scroll to the next sheet. Ctrl+Click to scroll to the last sheet. Right-click to see all sheets."
          onClick={(e) => scrollTabs(1, e.ctrlKey || e.metaKey)}
        >
          <ChevronRight size={14} strokeWidth={1.75} />
        </button>
      </div>
      <div
        ref={strip}
        className="sheet-strip"
        role="tablist"
        aria-label="Sheets"
        data-dragging={drag?.moved ? (drag.copy ? "copy" : "move") : undefined}
      >
        {sheets.map((sheet, index) =>
          sheet.hidden ? null : (
            <div
              key={index + ":" + sheet.name}
              data-sheet={index}
              role="tab"
              aria-selected={index === active}
              tabIndex={index === active ? 0 : -1}
              className="sheet-tab"
              data-grouped={index !== active && controller.group.has(index) ? "" : undefined}
              style={sheet.tabColor ? ({ "--tab-color": sheet.tabColor } as React.CSSProperties) : undefined}
              data-colored={sheet.tabColor ? (isDark(sheet.tabColor) ? "dark" : "light") : undefined}
              onPointerDown={(e) => onPointerDown(e, index)}
              onPointerMove={onPointerMove}
              onPointerUp={onPointerUp}
              onPointerCancel={() => {
                dragRef.current = null
                setDrag(null)
              }}
              onDoubleClick={() => commands.beginRename(index)}
              onContextMenu={(e) => {
                e.preventDefault()
                tabMenu(index, e.clientX, e.clientY)
              }}
            >
              {marker && drag?.before === index ? <span className="sheet-drop-marker" aria-hidden /> : null}
              {commands.renaming === index ? (
                <RenameBox controller={controller} index={index} />
              ) : (
                <span className="sheet-tab-name">{sheet.name}</span>
              )}
            </div>
          ),
        )}
        {marker && drag?.before === sheets.length ? <span className="sheet-drop-marker is-end" aria-hidden /> : null}
      </div>
      <button
        type="button"
        className="sheet-add"
        aria-label="New sheet"
        title="New sheet (Shift+F11)"
        disabled={commands.isWorkbookProtected()}
        onClick={() => commands.insert(active + 1)}
      >
        <Plus size={15} strokeWidth={1.75} />
      </button>
      {menu ? <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} /> : null}
      {colorAt ? (
        <TabColorPopup
          x={colorAt.x}
          y={colorAt.y}
          current={sheets[colorAt.index]?.tabColor}
          onPick={(hex) => {
            commands.setColor(colorAt.index, hex)
            setColorAt(null)
          }}
          onClose={() => setColorAt(null)}
        />
      ) : null}
    </div>
  )
}

/** Dark tab colours take white text (as Excel). */
function isDark(hex: string): boolean {
  const n = parseInt(hex.replace("#", ""), 16)
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255]
  return 0.299 * r + 0.587 * g + 0.114 * b < 140
}

/** The name being edited on the tab: Enter or a click elsewhere keeps it, Esc puts the old name back. */
function RenameBox({ controller, index }: { controller: SheetController; index: number }) {
  const name = controller.wb.sheets[index].name
  const [text, setText] = useState(name)
  const ref = useRef<HTMLInputElement>(null)
  const done = useRef(false)

  useEffect(() => {
    ref.current?.focus()
    ref.current?.select()
  }, [])

  const commit = () => {
    if (done.current) return
    const problem = controller.sheets.rename(index, text.trim() === "" ? "" : text)
    if (problem) {
      // Excel says why, then goes back to the name being typed.
      done.current = true
      controller.shell.alert(problem, () =>
        requestAnimationFrame(() => {
          done.current = false
          ref.current?.focus()
          ref.current?.select()
        }),
      )
      return
    }
    done.current = true
    controller.sheets.endRename()
  }

  return (
    <input
      ref={ref}
      className="sheet-rename"
      aria-label="Sheet name"
      value={text}
      maxLength={31}
      size={Math.max(4, text.length + 1)}
      spellCheck={false}
      data-own-escape=""
      onChange={(e) => setText(e.target.value)}
      onPointerDown={(e) => e.stopPropagation()}
      onBlur={() => commit()}
      onKeyDown={(e) => {
        e.stopPropagation()
        if (e.key === "Enter") {
          e.preventDefault()
          commit()
        } else if (e.key === "Escape") {
          e.preventDefault()
          done.current = true
          controller.sheets.endRename()
        }
      }}
    />
  )
}

function TabColorPopup({
  x,
  y,
  current,
  onPick,
  onClose,
}: {
  x: number
  y: number
  current?: string
  onPick: (hex: string | null) => void
  onClose: () => void
}) {
  const ref = useRef<HTMLDivElement>(null)
  const [place, setPlace] = useState({ left: x, top: y, visible: false })
  useLayoutEffect(() => {
    const rect = ref.current?.getBoundingClientRect()
    if (!rect) return
    setPlace({
      left: Math.max(4, Math.min(x, window.innerWidth - rect.width - 4)),
      top: Math.max(4, y - rect.height),
      visible: true,
    })
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
    return () => {
      window.removeEventListener("pointerdown", onDown, true)
      window.removeEventListener("keydown", onKey, true)
    }
  }, [onClose])
  return (
    <div
      ref={ref}
      className="context-menu tab-color-popup"
      role="menu"
      aria-label="Tab Color"
      style={{ left: place.left, top: place.top, visibility: place.visible ? "visible" : "hidden" }}
    >
      <ColorPanel noneLabel="No Color" current={current} onPick={onPick} />
    </div>
  )
}
