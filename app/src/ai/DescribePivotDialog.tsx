import type { PivotCache, PivotRecipe, PivotTable } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "@/shell/Dialog"
import { PivotPreview } from "@/shell/RecommendedPivotsDialog"
import { DescribePivot, useAiReady } from "./DescribePivot"

/**
 * "Describe the PivotTable you want" from the PivotTable Fields pane: the AI's layout for the PivotTable the pane is
 * on, previewed, and put in place only on Apply (one undo step; its name, place, style and options stay).
 */
export function DescribePivotDialog({
  controller: c,
  onClose,
  onOptions,
}: {
  controller: SheetController
  onClose: () => void
  onOptions: () => void
}) {
  const ready = useAiReady()
  const [at] = useState(() => c.pivots.activeWithCache())
  const [made, setMade] = useState<{ recipe: PivotRecipe; table: PivotTable; cache: PivotCache } | null>(null)
  if (!at) return null
  if (!ready)
    return (
      <Dialog title="Describe PivotTable" onCancel={onClose} onAccept={onOptions} acceptLabel="Options…" width={420}>
        <p className="dialog-note">
          AI is off. Turn it on in File › Options › AI, and choose which AI to use: one on this computer, or a service
          you have a key for.
        </p>
      </Dialog>
    )
  const apply = () => {
    if (!made) return
    onClose()
    c.pivots.applyRecipe(made.recipe)
  }
  return (
    <Dialog
      title="Describe PivotTable"
      onCancel={onClose}
      onAccept={apply}
      width={560}
      actions={
        <>
          <button type="button" className="dialog-btn is-primary" disabled={!made} onClick={apply}>
            Apply
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <DescribePivot
        cache={at.cache}
        autoFocus
        onRecipe={(recipe) => setMade({ recipe, ...c.pivots.recipePreview(at.cache, recipe) })}
      />
      <div className="recommend-preview describe-pivot-preview" aria-live="polite">
        {made ? (
          <>
            <div className="recommend-preview-title">{made.recipe.title}</div>
            <PivotPreview table={made.table} cache={made.cache} rows={20} />
          </>
        ) : (
          <p className="dialog-note">The PivotTable the AI proposes shows here. Apply puts it in place of this one.</p>
        )}
      </div>
    </Dialog>
  )
}
