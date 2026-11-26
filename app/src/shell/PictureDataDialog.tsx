import { ImageOff } from "lucide-react"
import { useMemo, useRef, useState } from "react"
import type { PictureGrid } from "@/pdf/picture"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

const key = (r: number, c: number) => `${r}:${c}`

/** A grid's rows, every row as wide as the widest. */
const widened = (g: PictureGrid) => {
  const width = Math.max(0, ...g.rows.map((r) => r.length))
  return g.rows.map((r) => Array.from({ length: width }, (_, k) => r[k] ?? ""))
}

/**
 * Data from Picture's review, as Excel's pane: the picture beside what was read in it, every cell editable, the ones
 * text recognition wasn't sure of marked. Next to check goes to the next marked cell; a marked cell counts as checked
 * once it's been visited. Insert Data puts the grid in at the active cell (asking first while cells are unchecked).
 */
export function PictureDataDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const review = controller.pictureData.review
  const [whole, setWhole] = useState(false)
  const source = review ? (whole && review.read.whole ? review.read.whole : review.read.table) : null
  const [rows, setRows] = useState<string[][]>(() => (source ? widened(source) : []))
  const [marked, setMarked] = useState<Set<string>>(() => new Set((source?.unsure ?? []).map(([r, c]) => key(r, c))))
  const [shown, setShown] = useState(true)
  const inputs = useRef(new Map<string, HTMLInputElement>())
  const order = useMemo(
    () =>
      [...marked].sort((a, b) => {
        const [ar, ac] = a.split(":").map(Number)
        const [br, bc] = b.split(":").map(Number)
        return ar - br || ac - bc
      }),
    [marked],
  )
  if (!review) return null

  const choose = (all: boolean) => {
    const g = all && review.read.whole ? review.read.whole : review.read.table
    setWhole(all)
    setRows(widened(g))
    setMarked(new Set(g.unsure.map(([r, c]) => key(r, c))))
  }
  const close = () => {
    controller.pictureData.close()
    onClose()
  }
  const insert = () => {
    const go = () => {
      onClose()
      controller.pictureData.insert(rows)
      controller.pictureData.close()
    }
    if (marked.size)
      controller.shell.confirm(
        `${marked.size === 1 ? "1 cell hasn't" : `${marked.size} cells haven't`} been checked. Insert the data anyway?`,
        go,
      )
    else go()
  }
  const next = () => {
    const first = order[0]
    if (first) inputs.current.get(first)?.focus()
  }
  const checked = (k: string) =>
    setMarked((m) => {
      if (!m.has(k)) return m
      const copy = new Set(m)
      copy.delete(k)
      return copy
    })

  return (
    <Dialog
      title="Data from Picture"
      onCancel={close}
      onAccept={insert}
      width={900}
      actions={
        <>
          <span className="qe-status picture-status">
            {marked.size ? `${marked.size} cell${marked.size === 1 ? "" : "s"} to check` : "Everything's checked"}
          </span>
          <button type="button" className="dialog-btn" onClick={next} disabled={!marked.size}>
            <u>N</u>ext to Check
          </button>
          <button type="button" className="dialog-btn is-primary" onClick={insert}>
            <u>I</u>nsert Data
          </button>
          <button type="button" className="dialog-btn" onClick={close}>
            Cancel
          </button>
        </>
      }
    >
      <div className="picture-dialog">
        <div className="picture-side">
          {shown ? (
            <img className="picture-image" src={review.url} alt="The picture" onError={() => setShown(false)} />
          ) : (
            <p className="picture-missing">
              <ImageOff size={28} aria-hidden />
              This kind of picture can&apos;t be shown here, but its text was read.
            </p>
          )}
          {review.read.whole ? (
            <div className="picture-choice" role="radiogroup" aria-label="What to insert">
              <label className="dialog-radio">
                <input type="radio" checked={!whole} onChange={() => choose(false)} />
                The table
              </label>
              <label className="dialog-radio">
                <input type="radio" checked={whole} onChange={() => choose(true)} />
                Everything in the picture
              </label>
            </div>
          ) : null}
          <p className="dialog-note">
            {marked.size ? (
              <>
                <span className="nav-unsure-swatch" aria-hidden /> Marked cells: text recognition wasn&apos;t sure of
                them. Check each against the picture and correct it here.
              </>
            ) : (
              "Correct anything that was read wrong, then insert the data at the selected cell."
            )}
          </p>
        </div>
        <div className="picture-pane">
          <table className="qe-grid picture-grid" aria-label="What was read">
            <tbody>
              {rows.map((r, i) => (
                <tr key={i}>
                  {r.map((v, k) => {
                    const id = key(i, k)
                    return (
                      <td key={k} className={marked.has(id) ? "is-unsure" : undefined}>
                        <input
                          ref={(el) => {
                            if (el) inputs.current.set(id, el)
                            else inputs.current.delete(id)
                          }}
                          aria-label={`Row ${i + 1}, column ${k + 1}${marked.has(id) ? ", to check" : ""}`}
                          title={marked.has(id) ? "Text recognition wasn't sure of this" : undefined}
                          // A value wrapped over lines keeps them unless it's edited (a field shows one line).
                          value={v.replace(/\n/g, " ")}
                          size={Math.max(4, Math.min(40, v.length + 1))}
                          onChange={(e) => {
                            const text = e.target.value
                            setRows((all) =>
                              all.map((row, ri) => (ri === i ? row.map((c, ci) => (ci === k ? text : c)) : row)),
                            )
                            checked(id)
                          }}
                          onBlur={() => checked(id)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") {
                              e.preventDefault()
                              e.stopPropagation()
                              checked(id)
                              const after = order.find((o) => o !== id)
                              if (after) inputs.current.get(after)?.focus()
                            }
                          }}
                        />
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </Dialog>
  )
}
