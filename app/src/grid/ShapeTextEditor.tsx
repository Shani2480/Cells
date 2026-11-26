import { useEffect, useRef } from "react"
import {
  colorCss,
  lookAt,
  sameTextLook,
  textLength,
  tidyText,
  type ShapeText,
  type SheetShape,
  type TextLook,
  type TextParagraph,
  type TextRun,
  type Theme,
} from "@cells/core"
import { shapeGeometry } from "@/drawing/path"
import type { SheetController } from "@/sheet/controller"

const PT = 96 / 72
const DEFAULT_INSETS = { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }
/** Where focus may go without closing the editor: the ribbon, its menus and dialogs (they format the selection). */
const KEEPS_EDITING = ".ribbon, .popover, .menu, [role='dialog'], [role='menu'], [role='listbox'], .qat"

/**
 * Typing formatted text into a shape, in place: the text shows in its own fonts, sizes and colours where the shape
 * draws it (the shape is drawn without its text meanwhile), and Home's font and alignment commands format what's
 * selected (Ctrl+B, I and U too). Typed text takes the look of the text it's typed into, as in Excel. The text is
 * kept when the editor closes (focus leaves it for anything but the ribbon and its menus, or Escape).
 */
export function ShapeTextEditor({
  shape,
  box,
  zoom,
  theme,
  controller,
  onDone,
}: {
  shape: SheetShape
  /** The shape's size (pixels at 100%). */
  box: { w: number; h: number }
  zoom: number
  theme: Theme
  controller: SheetController
  onDone: () => void
}) {
  const objects = controller.objects
  const el = useRef<HTMLDivElement>(null)
  const model = useRef<ShapeText>(startText(shape.text))
  const looks = useRef<TextLook[]>([])
  const saved = useRef<[number, number]>([0, 0])
  const done = useRef(false)
  const shown = useRef<TextLook | null>(null)
  const composing = useRef(false)

  const finish = () => {
    if (done.current) return
    done.current = true
    const text = tidyText(model.current)
    objects.lastCaret = { id: shape.id, fromEnd: textLength(text) - saved.current[1] }
    objects.caret = null
    objects.setRichText(shape.id, text)
    onDone()
  }
  const latest = useRef(finish)
  latest.current = finish

  useEffect(() => {
    const div = el.current!
    const paint = () => (looks.current = paintText(div, model.current, theme))
    const read = () => {
      model.current = readText(div, model.current, looks.current)
      // Text typed outside a paragraph (after everything was deleted): drawn again in the paragraphs' shape.
      if (!composing.current && !canonical(div)) {
        const sel = selectionIn(div) ?? saved.current
        paint()
        place(div, ...sel)
      }
    }
    const remember = () => {
      const sel = selectionIn(div)
      if (!sel) return
      saved.current = sel
      // The ribbon follows the look at the caret (only told when it changes).
      const look = lookAt(model.current, ...sel)
      if (!shown.current || !sameTextLook(shown.current, look)) {
        shown.current = look
        controller.changed()
      }
    }
    paint()
    div.focus()
    const caret = objects.caret
    if (caret?.id === shape.id) {
      const at = Math.max(0, textLength(model.current) - caret.fromEnd)
      place(div, at, at)
    } else place(div, 0, textLength(model.current))
    saved.current = selectionIn(div) ?? [0, 0]

    objects.textEditor = {
      text: () => model.current,
      range: () => selectionIn(div) ?? saved.current,
      set: (text) => {
        const sel = selectionIn(div) ?? saved.current
        model.current = text
        paint()
        div.focus()
        place(div, ...sel)
        saved.current = sel
        shown.current = null
        controller.changed()
      },
    }
    objects.textInserter = (insert) => {
      div.focus()
      place(div, ...saved.current)
      document.execCommand("insertText", false, insert)
    }
    objects.finishText = () => latest.current()
    const onInput = () => read()
    const onSelection = () => remember()
    // A click anywhere but the editor, the ribbon or its menus closes it (focus may not move on a click on the sheet).
    const onPointer = (e: PointerEvent) => {
      const t = e.target as Element | null
      if (t && (div.contains(t) || t.closest(KEEPS_EDITING))) return
      latest.current()
    }
    div.addEventListener("input", onInput)
    document.addEventListener("selectionchange", onSelection)
    document.addEventListener("pointerdown", onPointer, true)
    return () => {
      div.removeEventListener("input", onInput)
      document.removeEventListener("selectionchange", onSelection)
      document.removeEventListener("pointerdown", onPointer, true)
      objects.textEditor = null
      objects.textInserter = null
      objects.finishText = null
    }
    // The editor is set up once per shape; its text lives in `model` from then on.
  }, [shape.id])

  const text = model.current
  const ins = text.insets ?? DEFAULT_INSETS
  const g = shapeGeometry(shape, box.w, box.h).text
  const inner = {
    x: g.x + ins.l,
    y: g.y + ins.t,
    w: Math.max(8, g.w - ins.l - ins.r),
    h: Math.max(8, g.h - ins.t - ins.b),
  }
  const anchor = text.anchor === "ctr" ? "center" : text.anchor === "b" ? "flex-end" : "flex-start"
  return (
    <div
      className="shape-text-editor"
      style={{
        left: inner.x * zoom,
        top: inner.y * zoom,
        width: inner.w,
        height: inner.h,
        transform: `scale(${zoom})`,
        justifyContent: anchor,
      }}
      onPointerDown={(e) => e.stopPropagation()}
    >
      <div
        ref={el}
        className="shape-text-content"
        contentEditable
        suppressContentEditableWarning
        role="textbox"
        aria-multiline
        aria-label="Shape text"
        style={{ whiteSpace: text.wrap === false ? "pre" : "pre-wrap" }}
        onCompositionStart={() => (composing.current = true)}
        onCompositionEnd={() => (composing.current = false)}
        onBlur={(e) => {
          const to = e.relatedTarget as Element | null
          if (to && (to.closest(KEEPS_EDITING) || el.current?.contains(to))) return
          // Focus going nowhere in particular (a click on the sheet): the pointer handler decides.
          if (!to) return
          finish()
        }}
        onPaste={(e) => {
          // Pasted as plain text, in the look at the caret.
          e.preventDefault()
          document.execCommand("insertText", false, e.clipboardData.getData("text/plain").replace(/\r\n?/g, "\n"))
        }}
        onKeyDown={(e) => {
          e.stopPropagation()
          const mod = e.ctrlKey || e.metaKey
          if (e.key === "Escape") {
            e.preventDefault()
            finish()
          } else if (e.key === "Tab") {
            e.preventDefault()
            document.execCommand("insertText", false, "\t")
          } else if (e.key === "Enter" && e.shiftKey) {
            // One kind of break: a new paragraph.
            e.preventDefault()
            document.execCommand("insertParagraph")
          } else if (mod && !e.altKey && ["b", "i", "u", "5"].includes(e.key.toLowerCase())) {
            e.preventDefault()
            const k = e.key.toLowerCase()
            if (k === "u") controller.format.underline()
            else controller.format.toggle(k === "b" ? "bold" : k === "i" ? "italic" : "strike")
          }
        }}
      />
    </div>
  )
}

/** The text to start from: a shape without text gets one empty paragraph. */
function startText(text: ShapeText | undefined): ShapeText {
  if (!text || !text.paragraphs.length) return { ...(text ?? {}), paragraphs: [{ runs: [] }] }
  return text
}

/** A run's look as CSS. */
function runStyle(look: TextLook, theme: Theme): Partial<CSSStyleDeclaration> {
  const color = colorCss(look.color ?? { theme: "dk1" }, theme)
  const deco = [look.underline ? "underline" : "", look.strike ? "line-through" : ""].filter(Boolean).join(" ")
  return {
    fontFamily: `"${look.font ?? theme.fonts.minor}", "Aptos", "Calibri", sans-serif`,
    fontSize: `${(look.size ?? 11) * PT}px`,
    fontWeight: look.bold ? "700" : "400",
    fontStyle: look.italic ? "italic" : "normal",
    textDecoration: deco || "none",
    color:
      color.alpha < 1 ? `color-mix(in srgb, ${color.hex} ${Math.round(color.alpha * 100)}%, transparent)` : color.hex,
  }
}

const ALIGN: Record<NonNullable<TextParagraph["align"]>, string> = {
  l: "left",
  ctr: "center",
  r: "right",
  just: "justify",
}

/** The text drawn as the editor's DOM: a div per paragraph, a span per run (its look's index in `data-look`). */
function paintText(div: HTMLElement, text: ShapeText, theme: Theme): TextLook[] {
  const looks: TextLook[] = []
  const lookIndex = (run: TextRun | undefined) => {
    const { text: _t, ...look } = run ?? { text: "" }
    let i = looks.findIndex((l) => sameTextLook(l, look))
    if (i < 0) i = looks.push(look) - 1
    return i
  }
  div.replaceChildren(
    ...text.paragraphs.map((p) => {
      const para = document.createElement("div")
      para.className = "shape-text-p"
      if (p.align) para.style.textAlign = ALIGN[p.align]
      para.dataset.align = p.align ?? ""
      // The paragraph's own look (what's typed into it when it's empty): its first run's.
      const own = lookIndex(p.runs[0])
      para.dataset.look = String(own)
      const runs = p.runs.filter((r) => r.text)
      if (p.math) {
        para.contentEditable = "false"
        para.textContent = "⧼equation⧽"
      } else if (!runs.length) {
        const span = document.createElement("span")
        span.dataset.look = String(own)
        Object.assign(span.style, runStyle(looks[own], theme))
        span.appendChild(document.createElement("br"))
        para.appendChild(span)
      } else
        for (const r of runs) {
          const span = document.createElement("span")
          const i = lookIndex(r)
          span.dataset.look = String(i)
          Object.assign(span.style, runStyle(looks[i], theme))
          span.textContent = r.text
          para.appendChild(span)
        }
      return para
    }),
  )
  return looks
}

/** Whether the DOM is still a div per paragraph (typing after deleting everything can leave text outside them). */
function canonical(div: HTMLElement): boolean {
  return [...div.childNodes].every((n) => n instanceof HTMLElement && n.classList.contains("shape-text-p"))
}

/** The paragraphs' nodes: each paragraph div, and loose nodes between them as paragraphs of their own. */
function paragraphNodes(div: HTMLElement): Node[][] {
  const out: Node[][] = []
  let loose: Node[] | null = null
  for (const n of div.childNodes) {
    if (n instanceof HTMLElement && (n.tagName === "DIV" || n.tagName === "P")) {
      loose = null
      out.push([n])
    } else if (n instanceof HTMLElement && n.tagName === "BR") {
      loose = null
    } else {
      if (!loose) out.push((loose = []))
      loose.push(n)
    }
  }
  return out
}

/** The editor's DOM back as text: each text node in the look of the span it's in (or the one before it). */
function readText(div: HTMLElement, before: ShapeText, looks: TextLook[]): ShapeText {
  const lookOfNode = (node: Node, para: Node[]): TextLook | null => {
    for (let e = node.parentElement; e && e !== div; e = e.parentElement) {
      if (e.dataset.look !== undefined) return looks[Number(e.dataset.look)] ?? null
    }
    const own = para[0] instanceof HTMLElement ? para[0].dataset.look : undefined
    return own !== undefined ? (looks[Number(own)] ?? null) : null
  }
  const paragraphs: TextParagraph[] = paragraphNodes(div).map((nodes) => {
    const first = nodes[0]
    const para: TextParagraph = { runs: [] }
    if (first instanceof HTMLElement) {
      const align = first.dataset.align
      if (align === "l" || align === "ctr" || align === "r" || align === "just") para.align = align
    }
    let last: TextLook = looks[0] ?? {}
    const paraLook =
      first instanceof HTMLElement && first.dataset.look !== undefined ? looks[Number(first.dataset.look)] : undefined
    for (const root of nodes) {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      if (root.nodeType === Node.TEXT_NODE) {
        const look = paraLook ?? last
        para.runs.push({ ...look, text: root.textContent ?? "" })
        continue
      }
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        const look = lookOfNode(t, nodes) ?? last
        last = look
        para.runs.push({ ...look, text: t.textContent ?? "" })
      }
    }
    if (!para.runs.some((r) => r.text)) para.runs = [{ ...(paraLook ?? last), text: "" }]
    return para
  })
  // Equations stay as they were (shown as a placeholder the editor can't change).
  const maths = before.paragraphs.filter((p) => p.math)
  let m = 0
  const out = paragraphs.map((p) =>
    p.runs.length === 1 && p.runs[0].text === "⧼equation⧽" && maths[m] ? maths[m++] : p,
  )
  return tidyText({ ...before, paragraphs: out.length ? out : [{ runs: [{ ...(looks[0] ?? {}), text: "" }] }] })
}

/** The selection in the editor as positions in its plain text, or null when it's elsewhere. */
function selectionIn(div: HTMLElement): [number, number] | null {
  const sel = document.getSelection()
  if (!sel || !sel.rangeCount) return null
  const r = sel.getRangeAt(0)
  if (!div.contains(r.startContainer) || !div.contains(r.endContainer)) return null
  const a = positionOf(div, r.startContainer, r.startOffset)
  const b = positionOf(div, r.endContainer, r.endOffset)
  return a <= b ? [a, b] : [b, a]
}

function positionOf(div: HTMLElement, node: Node, offset: number): number {
  const paras = paragraphNodes(div)
  let at = 0
  for (const nodes of paras) {
    const start = nodes[0]
    const end = nodes[nodes.length - 1]
    const range = document.createRange()
    range.setStartBefore(start)
    range.setEndAfter(end)
    if (range.isPointInRange(node, offset) || nodes.some((n) => n === node || n.contains(node))) {
      const part = document.createRange()
      part.setStartBefore(start)
      part.setEnd(node, offset)
      return at + part.toString().length
    }
    at += range.toString().length + 1
  }
  return Math.max(0, at - 1)
}

/** Puts the selection at positions of the plain text. */
function place(div: HTMLElement, from: number, to: number) {
  const point = (pos: number): [Node, number] => {
    let at = 0
    const paras = paragraphNodes(div)
    for (let i = 0; i < paras.length; i++) {
      const nodes = paras[i]
      const texts: Text[] = []
      for (const root of nodes) {
        if (root.nodeType === Node.TEXT_NODE) texts.push(root as Text)
        const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
        for (let t = w.nextNode(); t; t = w.nextNode()) texts.push(t as Text)
      }
      const len = texts.reduce((n, t) => n + t.length, 0)
      if (pos <= at + len || i === paras.length - 1) {
        let rest = Math.min(pos - at, len)
        for (const t of texts) {
          if (rest <= t.length) return [t, rest]
          rest -= t.length
        }
        // An empty paragraph: inside its span, before the <br>.
        const first = nodes[0]
        const span = first instanceof HTMLElement ? (first.querySelector("span") ?? first) : first
        return [span, 0]
      }
      at += len + 1
    }
    return [div, 0]
  }
  const sel = document.getSelection()
  if (!sel) return
  const [a, ao] = point(from)
  const [b, bo] = point(to)
  const range = document.createRange()
  range.setStart(a, ao)
  range.setEnd(b, bo)
  sel.removeAllRanges()
  sel.addRange(range)
}
