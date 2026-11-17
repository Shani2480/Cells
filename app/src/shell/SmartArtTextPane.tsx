import {
  addSmartArtNode,
  demoteSmartArtNode,
  promoteSmartArtNode,
  removeSmartArtNode,
  setSmartArtText,
  smartArtLines,
  type SmartArtPath,
} from "@cells/core"
import { X } from "lucide-react"
import { useEffect, useRef, useState } from "react"
import type { SheetController } from "@/sheet/controller"

const same = (a: SmartArtPath | null, b: SmartArtPath) => !!a && a.length === b.length && a.every((v, i) => v === b[i])

/**
 * SmartArt Design › Text Pane: the SmartArt's text as a bulleted list, a line per point. Typing shows in the SmartArt
 * as you go and is one undo step per line. Enter adds a point after, Tab and Shift+Tab move it down or up a level,
 * Backspace on an empty line removes it, arrows go between lines.
 */
export function SmartArtTextPane({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  const objects = controller.objects
  const sa = objects.smartArt()
  const list = useRef<HTMLUListElement>(null)
  const [draft, setDraft] = useState<{ path: SmartArtPath; text: string } | null>(null)
  const point = objects.smartArtPoint
  // Opened on a point (a SmartArt just put in): its line takes the keyboard.
  useEffect(() => {
    const at = objects.smartArtPoint
    if (at) list.current?.querySelector<HTMLInputElement>(`[data-path="${at.join(".")}"]`)?.focus()
  }, [])
  // The current line keeps the keyboard when the tree changes under it.
  useEffect(() => {
    if (!point) return
    const input = list.current?.querySelector<HTMLInputElement>(`[data-path="${point.join(".")}"]`)
    if (input && document.activeElement !== input && list.current?.contains(document.activeElement)) input.focus()
  })
  if (!sa) return null
  const lines = smartArtLines(sa.smartArt.nodes)
  const close = () => {
    objects.setTextPane(false)
    onDone()
  }
  /** The line being typed in, as one undo step. */
  const commit = () => {
    if (!draft) return
    const at = lines.find((l) => same(draft.path, l.path))
    setDraft(null)
    if (at && at.node.text !== draft.text)
      objects.setSmartArt(sa.id, { nodes: setSmartArtText(sa.smartArt.nodes, draft.path, draft.text) }, "Typing")
    else objects.previewSmartArt(sa.id, null)
  }
  const go = (path: SmartArtPath) => {
    commit()
    objects.smartArtPoint = path
    controller.changed()
  }
  return (
    <aside
      className="comments-pane smartart-pane"
      aria-label="Text Pane"
      onKeyDown={(e) => {
        if (e.key === "Escape") {
          e.preventDefault()
          commit()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Type your text here</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <ul className="smartart-lines" ref={list}>
        {lines.map(({ node, path }, i) => {
          const editing = draft && same(draft.path, path)
          return (
            <li key={path.join(".")} style={{ paddingLeft: 8 + (path.length - 1) * 18 }}>
              <span className="smartart-bullet" aria-hidden>
                •
              </span>
              <input
                className="smartart-line"
                data-path={path.join(".")}
                aria-label={`Point ${i + 1}`}
                placeholder="[Text]"
                value={editing ? draft.text : node.text}
                onFocus={() => {
                  if (!same(objects.smartArtPoint, path)) go(path)
                }}
                onBlur={commit}
                onChange={(e) => {
                  const text = e.target.value
                  setDraft({ path, text })
                  objects.previewSmartArt(sa.id, setSmartArtText(sa.smartArt.nodes, path, text))
                }}
                onKeyDown={(e) => {
                  const run = (edit: Parameters<typeof objects.editSmartArt>[0], label: string) => {
                    e.preventDefault()
                    commit()
                    objects.smartArtPoint = path
                    objects.editSmartArt(edit, label)
                  }
                  if (e.key === "Enter") run((n, p) => addSmartArtNode(n, p), "Add Shape")
                  else if (e.key === "Tab" && e.shiftKey) run((n, p) => p && promoteSmartArtNode(n, p), "Promote")
                  else if (e.key === "Tab") run((n, p) => p && demoteSmartArtNode(n, p), "Demote")
                  else if (e.key === "Backspace" && !e.currentTarget.value && lines.length > 1)
                    run((n, p) => p && removeSmartArtNode(n, p), "Delete")
                  else if (e.key === "ArrowUp" && i > 0) {
                    e.preventDefault()
                    go(lines[i - 1]!.path)
                  } else if (e.key === "ArrowDown" && i < lines.length - 1) {
                    e.preventDefault()
                    go(lines[i + 1]!.path)
                  }
                }}
              />
            </li>
          )
        })}
      </ul>
      <p className="smartart-pane-name">{sa.name}</p>
    </aside>
  )
}
