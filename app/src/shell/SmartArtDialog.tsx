import { SMARTART_LAYOUTS, type SheetGroup, type SmartArtCategory, type SmartArtLayout, type Theme } from "@cells/core"
import { useState } from "react"
import { objectSvg } from "@/drawing/render"
import { layoutSmartArt } from "@/drawing/smartart"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

const CATEGORIES: { id: SmartArtCategory | "all"; label: string }[] = [
  { id: "all", label: "All" },
  { id: "list", label: "List" },
  { id: "process", label: "Process" },
  { id: "cycle", label: "Cycle" },
  { id: "matrix", label: "Matrix" },
]

/** A layout drawn as a picture of itself, its shapes without text (`text`: with Office's "[Text]"). */
export function SmartArtPicture({
  layout,
  theme,
  w,
  h,
  colors,
  text = false,
}: {
  layout: SmartArtLayout
  theme: Theme
  w: number
  h: number
  colors?: string
  text?: boolean
}) {
  const smartArt = {
    layout: layout.id,
    nodes: Array.from({ length: layout.starts }, () => ({ text: "", children: [] })),
    ...(colors ? { colors } : {}),
  }
  const items = layoutSmartArt(smartArt, { w, h }, theme).map((s) => (text ? s : { ...s, text: undefined }))
  const group: SheetGroup = {
    kind: "group",
    id: 0,
    name: "",
    anchor: { from: { row: 0, col: 0, x: 0, y: 0 }, to: { row: 0, col: 0, x: 0, y: 0 } },
    child: { w, h },
    items,
  }
  return (
    <svg
      width={w}
      height={h}
      viewBox={`0 0 ${w} ${h}`}
      aria-hidden
      dangerouslySetInnerHTML={{ __html: objectSvg(group, w, h, theme) }}
    />
  )
}

/**
 * Insert › SmartArt (Choose a SmartArt Graphic): the layouts by kind on the left, their pictures in the middle, the
 * chosen one large with its name and what it's for on the right. OK (or a double-click) puts it in, ready to type
 * in its Text Pane.
 */
export function SmartArtDialog({
  controller,
  onClose,
  onInserted,
}: {
  controller: SheetController
  onClose: () => void
  /** Closes the dialog leaving the keyboard to the Text Pane. */
  onInserted: () => void
}) {
  const [category, setCategory] = useState<SmartArtCategory | "all">("all")
  const [picked, setPicked] = useState(SMARTART_LAYOUTS[0]!)
  const theme = controller.wb.theme
  const shown = SMARTART_LAYOUTS.filter((l) => category === "all" || l.category === category)
  const insert = (layout: SmartArtLayout) => {
    controller.objects.insertSmartArt(layout.id)
    onInserted()
  }
  return (
    <Dialog title="Choose a SmartArt Graphic" width={720} onCancel={onClose} onAccept={() => insert(picked)}>
      <div className="smartart-chooser">
        <ul className="smartart-kinds" role="listbox" aria-label="Kinds">
          {CATEGORIES.map((c) => (
            <li key={c.id}>
              <button
                type="button"
                role="option"
                aria-selected={category === c.id}
                className={`smartart-kind${category === c.id ? " is-selected" : ""}`}
                onClick={() => setCategory(c.id)}
              >
                {c.label}
              </button>
            </li>
          ))}
        </ul>
        <div className="smartart-layouts" role="listbox" aria-label="Layouts">
          {shown.map((l) => (
            <button
              key={l.id}
              type="button"
              role="option"
              aria-selected={picked.id === l.id}
              aria-label={l.name}
              title={l.name}
              data-layout={l.key}
              className={`smartart-layout${picked.id === l.id ? " is-selected" : ""}`}
              onClick={() => setPicked(l)}
              onDoubleClick={() => insert(l)}
            >
              <SmartArtPicture layout={l} theme={theme} w={84} h={64} />
            </button>
          ))}
        </div>
        <div className="smartart-about">
          <div className="smartart-about-picture">
            <SmartArtPicture layout={picked} theme={theme} w={200} h={150} />
          </div>
          <h3>{picked.name}</h3>
          <p>{picked.description}</p>
        </div>
      </div>
    </Dialog>
  )
}
