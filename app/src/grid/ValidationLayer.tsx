import { useEffect, useRef, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { colX, rowY, type GridLayout } from "./layout"

/** Excel's in-cell list button: 17 px square at 100%, just right of the cell, level with its bottom. */
const BUTTON = 17
/** Rows the in-cell list shows before it scrolls (Excel shows eight). */
const LIST_ROWS = 8
const ITEM_HEIGHT = 20

/**
 * Data Validation on the sheet: the in-cell list's button and list beside the active cell, the active cell's input
 * message, and the red circles of Circle Invalid Data.
 */
export function ValidationLayer({
  controller,
  layout,
  zoom,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  /** Back to the sheet (focus the grid's input). */
  onDone: () => void
}) {
  const sheet = controller.sheet
  const { row, col } = controller.selection.active
  // `pick`: Pick From Drop-down List (the column's own entries) rather than the cell's list.
  const [open, setOpen] = useState<{ row: number; col: number; pick?: boolean } | null>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const [pick, setPick] = useState(0)

  const visible = (r: number, c: number) =>
    layout.rows.some((l) => l.index === r) && layout.cols.some((l) => l.index === c)
  const rect = (r: number, c: number) => {
    const x = colX(sheet, layout, c)
    const y = rowY(sheet, layout, r)
    return { x, y, w: sheet.cols.size(c), h: sheet.rows.size(r) }
  }

  const own = controller.cellList(row, col)
  const picking = open?.pick === true && open.row === row && open.col === col
  const list = own ?? (picking ? controller.pickList(row, col) : null)
  const items = list?.items ?? null
  const editing = Boolean(controller.edit)
  const shown = visible(row, col) ? rect(row, col) : null
  // The list belongs to the cell it was opened on; moving away closes it.
  const isOpen = open !== null && open.row === row && open.col === col && items !== null && shown !== null

  const show = (pickList = false) => {
    const shownList = own ?? (pickList ? controller.pickList(row, col) : null)
    if (!shownList) return
    const at = shownList.items.findIndex((i) => i === shownList.current)
    setPick(Math.max(0, at))
    setOpen({ row, col, pick: !own && pickList })
  }

  useEffect(() => {
    const onOpen = (event: Event) => show(Boolean((event as CustomEvent<{ pick?: boolean }>).detail?.pick))
    window.addEventListener("cells:open-cell-list", onOpen)
    return () => window.removeEventListener("cells:open-cell-list", onOpen)
  })

  useEffect(() => {
    if (isOpen) listRef.current?.focus({ preventScroll: true })
  }, [isOpen])

  useEffect(() => {
    listRef.current?.querySelector<HTMLElement>(`[data-index="${pick}"]`)?.scrollIntoView({ block: "nearest" })
  }, [pick, isOpen])

  const close = () => {
    setOpen(null)
    onDone()
  }
  const choose = (index: number) => {
    if (!items || items[index] === undefined) return
    setOpen(null)
    list?.pick(items[index])
    onDone()
  }

  const message = !isOpen ? controller.validation.inputMessage() : null
  const circles = controller.validation.circled().filter((r) => visible(r.r0, r.c0))

  return (
    <>
      {circles.map((r) => {
        const c = rect(r.r0, r.c0)
        return (
          <div
            key={`${r.r0}:${r.c0}`}
            className="dv-circle"
            aria-hidden
            style={{
              left: (c.x - 4) * zoom,
              top: (c.y - 3) * zoom,
              width: (c.w + 8) * zoom,
              height: (c.h + 6) * zoom,
            }}
          />
        )
      })}
      {own && shown && !editing ? (
        <button
          type="button"
          className={`filter-button dv-list-button${isOpen ? " is-open" : ""}`}
          aria-label="Choose from the list"
          aria-haspopup="listbox"
          aria-expanded={isOpen}
          tabIndex={-1}
          style={{
            left: (shown.x + shown.w + 1) * zoom,
            top: (shown.y + shown.h - Math.min(BUTTON, shown.h)) * zoom,
            width: BUTTON * zoom,
            height: Math.min(BUTTON, shown.h) * zoom,
          }}
          onPointerDown={(event) => {
            event.stopPropagation()
            event.preventDefault()
          }}
          onClick={() => (isOpen ? close() : show())}
        >
          <svg width={7 * zoom} height={4 * zoom} viewBox="0 0 7 4" aria-hidden>
            <path d="M0 0h7L3.5 4z" fill="currentColor" />
          </svg>
        </button>
      ) : null}
      {isOpen && items && shown ? (
        <div
          ref={listRef}
          className="dv-list"
          role="listbox"
          tabIndex={-1}
          aria-label="List"
          aria-activedescendant={`dv-item-${pick}`}
          style={{
            left: shown.x * zoom,
            top: (shown.y + shown.h + 1) * zoom,
            minWidth: (shown.w + BUTTON + 1) * zoom,
            maxHeight: LIST_ROWS * ITEM_HEIGHT + 2,
          }}
          onPointerDown={(event) => event.stopPropagation()}
          onBlur={(event) => {
            if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setOpen(null)
          }}
          onKeyDown={(event) => {
            event.stopPropagation()
            const last = items.length - 1
            if (event.key === "ArrowDown") setPick(Math.min(last, pick + 1))
            else if (event.key === "ArrowUp") setPick(Math.max(0, pick - 1))
            else if (event.key === "PageDown") setPick(Math.min(last, pick + LIST_ROWS - 1))
            else if (event.key === "PageUp") setPick(Math.max(0, pick - LIST_ROWS + 1))
            else if (event.key === "Home") setPick(0)
            else if (event.key === "End") setPick(last)
            else if (event.key === "Enter" || event.key === "Tab") choose(pick)
            else if (event.key === "Escape") close()
            else return
            event.preventDefault()
          }}
        >
          {items.map((item, i) => (
            <div
              key={i}
              id={`dv-item-${i}`}
              data-index={i}
              role="option"
              aria-selected={i === pick}
              className="dv-item"
              style={{ height: ITEM_HEIGHT }}
              onPointerEnter={() => setPick(i)}
              onClick={() => choose(i)}
            >
              {item || " "}
            </div>
          ))}
        </div>
      ) : null}
      {message && shown ? (
        <div
          className="dv-input-message"
          role="tooltip"
          style={{ left: (shown.x + shown.w / 2) * zoom, top: (shown.y + shown.h + 6) * zoom }}
        >
          {message.title ? <div className="dv-input-title">{message.title}</div> : null}
          {message.prompt ? <div>{message.prompt}</div> : null}
        </div>
      ) : null}
    </>
  )
}
