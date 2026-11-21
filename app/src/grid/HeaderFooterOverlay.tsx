import { useEffect, useRef } from "react"
import type { SheetController } from "@/sheet/controller"
import type { HFSection } from "@/sheet/header-footer"
import type { GridLayout } from "./layout"
import { visiblePages, type PageRect } from "./page-layout"
import { CELL_FONT_FAMILY } from "./theme"

const SECTIONS: HFSection[] = ["left", "center", "right"]
const SECTION_NAMES: Record<HFSection, string> = { left: "Left", center: "Center", right: "Right" }

/** Where a page's header or footer sections sit (sheet pixels, as the layout's), with the text's size. */
export function headerFooterBox(controller: SheetController, page: PageRect, where: "header" | "footer") {
  const setup = controller.sheet.pageSetup
  const scale = controller.print.current().scale || 1
  const pxPerIn = 96 / scale
  const size = 11 / 0.75
  const min = size * 1.2 + 8
  if (where === "header") {
    const top = page.y + setup.margins.header * pxPerIn - 4
    return { x: page.body.x, y: top, w: page.body.w, h: Math.max(min, page.body.y - top - 2), size }
  }
  const bottom = page.y + page.h - setup.margins.footer * pxPerIn + 6
  const top = Math.min(page.body.y + page.body.h + 2, bottom - min)
  return { x: page.body.x, y: top, w: page.body.w, h: bottom - top, size }
}

/**
 * Page Layout: the header or footer being edited, as Excel shows it on the page: three boxes (left, center, right)
 * in the page's margin. Tab moves to the next box, Esc drops the changes; a click on the cells keeps them.
 */
export function HeaderFooterOverlay({
  controller,
  layout,
  zoom,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  onDone: () => void
}) {
  const editor = controller.headerFooter
  const e = editor.editing
  const refs = useRef<Partial<Record<HFSection, HTMLTextAreaElement | null>>>({})
  const page = e ? visiblePages(layout).find((p) => p.rowBand === e.rowBand && p.colBand === e.colBand) : undefined
  const focusKey = e ? `${e.where}:${e.rowBand}:${e.colBand}:${e.focus}` : ""

  // Opening, Go to Header/Footer and the element buttons put the caret back in the section being edited.
  useEffect(() => {
    const now = controller.headerFooter.editing
    if (!now) return
    const el = refs.current[now.section]
    if (!el) return
    el.focus({ preventScroll: true })
    const at = Math.min(now.caret[now.section], el.value.length)
    el.setSelectionRange(at, at)
  }, [controller, focusKey, page !== undefined])

  if (!e || !page) return null
  const box = headerFooterBox(controller, page, e.where)
  const third = box.w / 3
  return (
    <div
      className="hf-edit"
      role="group"
      aria-label={e.where === "header" ? "Header" : "Footer"}
      style={{ left: box.x * zoom, top: box.y * zoom, width: box.w * zoom, height: box.h * zoom }}
      onPointerDown={(event) => event.stopPropagation()}
    >
      {SECTIONS.map((s, i) => (
        <textarea
          key={s}
          ref={(el) => {
            refs.current[s] = el
          }}
          className={`hf-edit-section is-${s}`}
          aria-label={`${SECTION_NAMES[s]} ${e.where === "header" ? "header" : "footer"}`}
          spellCheck={false}
          value={e.text[s]}
          style={{
            left: i * third * zoom,
            width: third * zoom,
            fontSize: box.size * zoom,
            fontFamily: CELL_FONT_FAMILY,
            textAlign: s,
          }}
          onFocus={(event) => editor.place(s, event.currentTarget.selectionStart)}
          onSelect={(event) => editor.place(s, event.currentTarget.selectionStart)}
          onChange={(event) => editor.type(s, event.currentTarget.value, event.currentTarget.selectionStart)}
          onKeyDown={(event) => {
            if (event.key === "Escape") {
              event.preventDefault()
              editor.cancel()
              onDone()
            } else if (event.key === "Tab") {
              event.preventDefault()
              const next = SECTIONS[(i + (event.shiftKey ? 2 : 1)) % 3]
              refs.current[next]?.focus()
            }
          }}
        />
      ))}
    </div>
  )
}
