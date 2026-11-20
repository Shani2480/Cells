import {
  displayText,
  pivotLayout,
  type PivotCache,
  type PivotCell,
  type PivotRecipe,
  type PivotTable,
} from "@cells/core"
import { Sparkles } from "lucide-react"
import { useMemo, useState } from "react"
import { DescribePivot, useAiReady } from "@/ai/DescribePivot"
import type { SheetController } from "@/sheet/controller"
import { cn } from "@/lib/cn"
import { Dialog } from "./Dialog"

/**
 * Recommended PivotTables (Insert tab, PivotTable Analyze › Tools, Quick Analysis' Tables tab › More), as Excel's:
 * the suggestions down the left, each a small picture of its PivotTable; the one picked shown large on the right.
 * OK makes it on a new sheet; Blank PivotTable makes an empty one there for the Fields pane. With AI on, "Describe the
 * PivotTable you want" puts the AI's layout at the top of the list, picked, to look at before OK.
 */
export function RecommendedPivotsDialog({
  controller,
  source,
  onClose,
}: {
  controller: SheetController
  /** The data to suggest from (Quick Analysis gives its range); else the data at the active cell. */
  source?: string
  onClose: () => void
}) {
  const found = useMemo(() => controller.pivots.recommendations(source), [controller, source])
  const [picked, setPicked] = useState(0)
  const [described, setDescribed] = useState<{ recipe: PivotRecipe; table: PivotTable; cache: PivotCache } | null>(null)
  const aiReady = useAiReady()
  // With nothing to recommend, the list still opens when the person can describe one.
  if (typeof found === "string" || (!found.items.length && !aiReady)) {
    return (
      <Dialog title="Recommended PivotTables" onCancel={onClose} onAccept={onClose} cancelLabel={null}>
        <p className="dialog-label">
          {typeof found === "string" ? found : "Cells has no PivotTables to recommend for this data."}
        </p>
      </Dialog>
    )
  }
  const items = described ? [described, ...found.items] : found.items
  const make = (blank: boolean) => {
    const recipe = blank ? undefined : items[picked]?.recipe
    if (!blank && !recipe) return
    const problem = controller.pivots.create(found.source, { kind: "newSheet" }, undefined, recipe)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }
  const current = items[picked]
  return (
    <Dialog
      title="Recommended PivotTables"
      onCancel={onClose}
      onAccept={() => make(false)}
      width={760}
      actions={
        <>
          <button type="button" className="dialog-btn recommend-blank" onClick={() => make(true)}>
            Blank PivotTable
          </button>
          <button type="button" className="dialog-btn is-primary" disabled={!current} onClick={() => make(false)}>
            OK
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <DescribePivot
        cache={found.cache}
        onRecipe={(recipe) => {
          setDescribed({ recipe, ...controller.pivots.recipePreview(found.cache, recipe) })
          setPicked(0)
        }}
      />
      <div className="recommend-pivots">
        <div className="recommend-list" role="listbox" aria-label="Recommended PivotTables">
          {items.map((item, i) => (
            <button
              key={`${item === described ? "ai" : "r"}:${item.recipe.title}`}
              type="button"
              role="option"
              aria-selected={i === picked}
              className={cn("recommend-card", i === picked && "is-picked")}
              onClick={() => setPicked(i)}
              onDoubleClick={() => {
                setPicked(i)
                make(false)
              }}
              onKeyDown={(e) => {
                const step = e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0
                if (!step) return
                e.preventDefault()
                const next = Math.max(0, Math.min(items.length - 1, i + step))
                setPicked(next)
                ;(e.currentTarget.parentElement?.children[next] as HTMLElement | undefined)?.focus()
              }}
            >
              <PivotPreview table={item.table} cache={item.cache} rows={7} small />
              <span className="recommend-title">
                {item === described ? <Sparkles size={12} strokeWidth={1.75} aria-label="Described" /> : null}
                {item.recipe.title}
              </span>
            </button>
          ))}
        </div>
        <div className="recommend-preview" aria-live="polite">
          {current ? (
            <>
              <div className="recommend-preview-title">{current.recipe.title}</div>
              <PivotPreview table={current.table} cache={current.cache} rows={24} />
            </>
          ) : (
            <p className="dialog-note">Cells has no PivotTables to recommend for this data.</p>
          )}
        </div>
      </div>
    </Dialog>
  )
}

const BOLD = new Set<PivotCell["role"]>(["corner", "rowHeader", "colHeader", "dataHeader", "grandLabel", "grandValue"])

/** A PivotTable drawn as a plain grid of its cells (the first `rows` rows), for the dialog's pictures. */
export function PivotPreview({
  table,
  cache,
  rows,
  small,
}: {
  table: PivotTable
  cache: PivotCache
  rows: number
  small?: boolean
}) {
  const layout = useMemo(() => pivotLayout(table, cache), [table, cache])
  const { range } = layout
  const r1 = Math.min(range.r1, range.r0 + rows - 1)
  const at = new Map(layout.cells.map((c) => [`${c.row},${c.col}`, c]))
  const lines = []
  for (let r = range.r0; r <= r1; r++) {
    const cells = []
    for (let c = range.c0; c <= range.c1; c++) {
      const cell = at.get(`${r},${c}`)
      const v = cell?.value
      const text = v === undefined || v === null ? "" : displayText(v, cell?.numFmt ?? "General")
      cells.push(
        <td
          key={c}
          className={cn(
            typeof v === "number" && "is-number",
            cell && BOLD.has(cell.role) && "is-bold",
            cell && (cell.role === "grandLabel" || cell.role === "grandValue") && "is-grand",
            cell &&
              (cell.role === "corner" ||
                cell.role === "rowHeader" ||
                cell.role === "colHeader" ||
                cell.role === "dataHeader") &&
              "is-head",
          )}
          style={cell?.indent ? { paddingLeft: 4 + cell.indent * (small ? 6 : 10) } : undefined}
        >
          {text}
        </td>,
      )
    }
    lines.push(<tr key={r}>{cells}</tr>)
  }
  return (
    <table className={cn("pivot-preview", small && "is-small")} aria-hidden={small}>
      <tbody>{lines}</tbody>
    </table>
  )
}
