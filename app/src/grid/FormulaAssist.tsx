import { functionAtCaret } from "@cells/core"
import type { CSSProperties } from "react"
import type { SheetController } from "@/sheet/controller"

/**
 * The formula being typed with its references in their colours (the same colours outline the cells on the grid).
 * It sits under the editor, whose own text is transparent, so caret and selection stay the browser's.
 */
export function FormulaHighlight({
  controller,
  className,
  style,
  scrollLeft = 0,
}: {
  controller: SheetController
  className: string
  style?: CSSProperties
  scrollLeft?: number
}) {
  const edit = controller.edit
  if (!edit || !edit.text.startsWith("=")) return null
  const refs = controller.editReferences()
  const parts: React.ReactNode[] = []
  let at = 0
  for (const ref of refs) {
    if (ref.start > at) parts.push(edit.text.slice(at, ref.start))
    parts.push(
      <span key={ref.start} className="formula-ref" style={{ color: `var(--ref-${ref.color})` }}>
        {edit.text.slice(ref.start, ref.end)}
      </span>,
    )
    at = ref.end
  }
  parts.push(edit.text.slice(at))
  return (
    <div className={className} style={style} aria-hidden>
      <div style={{ transform: `translateX(${-scrollLeft}px)` }}>
        {parts}
        {/* Keeps a trailing line break's height. */}
        {"​"}
      </div>
    </div>
  )
}

/**
 * Under the editor: the list of functions matching the name being typed (Tab inserts the highlighted one), or the
 * argument tooltip of the function the caret is in, with the current argument in bold.
 */
export function FormulaAssist({
  controller,
  style,
  onPick,
}: {
  controller: SheetController
  style: CSSProperties
  onPick: () => void
}) {
  const edit = controller.edit
  if (!edit || !edit.text.startsWith("=")) return null
  const list = controller.autocomplete
  if (list) {
    const picked = list.items[list.pick]
    return (
      <div className="formula-assist" style={style}>
        <ul className="formula-list" role="listbox" aria-label="Functions">
          {list.items.map((info, i) => (
            <li
              key={info.name}
              role="option"
              aria-selected={i === list.pick}
              className={i === list.pick ? "is-picked" : undefined}
              onPointerDown={(event) => event.preventDefault()}
              onDoubleClick={() => {
                controller.acceptAutocomplete(i)
                onPick()
              }}
            >
              <span className="formula-list-icon" aria-hidden>
                fx
              </span>
              {info.name}
            </li>
          ))}
        </ul>
        {picked ? <div className="formula-list-note">{picked.description}</div> : null}
      </div>
    )
  }
  const at = functionAtCaret(edit.text, edit.selEnd)
  if (!at) return null
  const { info, arg } = at
  // Past the last named argument of a "…" list, the repeating one stays in bold.
  const repeating = info.args.at(-1) === "..."
  const bold = repeating && arg >= info.args.length - 1 ? info.args.length - 2 : arg
  return (
    <div className="formula-assist" style={style}>
      <div className="formula-tip" role="tooltip">
        <span className="formula-tip-name">{info.name}</span>(
        {info.args.map((name, i) => (
          <span key={i}>
            {i > 0 ? ", " : ""}
            <span className={i === bold ? "is-current" : undefined}>{name}</span>
          </span>
        ))}
        )
      </div>
    </div>
  )
}
