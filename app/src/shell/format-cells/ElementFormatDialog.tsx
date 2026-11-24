import { useState } from "react"
import type { CellStyle, StylePatch, TableStyleElement, Theme } from "@cells/core"
import { Dialog } from "../Dialog"
import type { BorderPlan } from "./borders"
import { BorderTab, FillTab, FontTab } from "./FormatCellsDialog"

type Tab = "font" | "border" | "fill"

/** A style element as the Format Cells tabs show it: font and fill as a cell style, borders as a border plan. */
function asCell(e: TableStyleElement): CellStyle {
  const out: CellStyle = {}
  if (e.bold) out.bold = true
  if (e.italic) out.italic = true
  if (e.strike) out.strike = true
  if (e.underline && e.underline !== "none") out.underline = e.underline
  if (e.color) out.fontColor = e.color
  if (e.fill) out.fill = e.fill
  return out
}

const asPlan = (e: TableStyleElement): BorderPlan => {
  const b = e.border ?? {}
  return { top: b.top, bottom: b.bottom, left: b.left, right: b.right, insideH: b.horizontal, insideV: b.vertical }
}

/** The element back from the tabs: font flags only when on (as Excel's dialog writes them), fill, borders. */
function fromTabs(base: TableStyleElement, style: CellStyle, plan: BorderPlan): TableStyleElement {
  const out: TableStyleElement = {}
  if (style.bold) out.bold = true
  if (style.italic) out.italic = true
  if (style.strike) out.strike = true
  if (style.underline) out.underline = style.underline
  if (style.fontColor) out.color = style.fontColor
  if (style.fill) out.fill = style.fill
  const border: NonNullable<TableStyleElement["border"]> = {}
  const sides = [
    ["top", "top"],
    ["bottom", "bottom"],
    ["left", "left"],
    ["right", "right"],
    ["insideH", "horizontal"],
    ["insideV", "vertical"],
  ] as const
  for (const [edge, side] of sides) {
    const b = plan[edge]
    if (b) border[side] = b
  }
  if (Object.keys(border).length) out.border = border
  if (base.stripe) out.stripe = base.stripe
  return out
}

/**
 * Format Cells for one element of a table or PivotTable style (New PivotTable Style › Format): Font (no font name or
 * size, as Excel), Border (with inside lines) and Fill; Clear empties all three.
 */
export function ElementFormatDialog({
  element,
  theme,
  onAccept,
  onClose,
}: {
  element: TableStyleElement
  theme?: Theme
  onAccept: (element: TableStyleElement) => void
  onClose: () => void
}) {
  const [tab, setTab] = useState<Tab>("font")
  const [style, setStyle] = useState<CellStyle>(() => asCell(element))
  const [plan, setPlan] = useState<BorderPlan>(() => asPlan(element))
  const set = (patch: StylePatch) =>
    setStyle((s) => {
      const next = { ...s } as Record<string, unknown>
      for (const [k, v] of Object.entries(patch)) {
        if (v === null || v === undefined) delete next[k]
        else next[k] = v
      }
      return next as CellStyle
    })
  const accept = () => {
    onAccept(fromTabs(element, style, plan))
    onClose()
  }
  return (
    <Dialog
      title="Format Cells"
      onCancel={onClose}
      onAccept={accept}
      width={560}
      actions={
        <>
          <button
            type="button"
            className="dialog-btn dialog-btn-start"
            onClick={() => {
              setStyle({})
              setPlan({})
            }}
          >
            Clear
          </button>
          <button type="button" className="dialog-btn is-primary" onClick={accept}>
            OK
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <div className="fc-tabs" role="tablist">
        {(
          [
            ["font", "Font"],
            ["border", "Border"],
            ["fill", "Fill"],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            type="button"
            role="tab"
            aria-selected={tab === id}
            className="fc-tab"
            onClick={() => setTab(id)}
          >
            {label}
          </button>
        ))}
      </div>
      <div className="fc-page" role="tabpanel" aria-label={tab}>
        {tab === "font" ? <FontTab style={style} set={set} theme={theme} differential /> : null}
        {tab === "border" ? (
          <BorderTab
            plan={plan}
            multiRow
            multiCol
            noDiagonals
            onChange={(edges) => setPlan((p) => ({ ...p, ...edges }))}
            theme={theme}
          />
        ) : null}
        {tab === "fill" ? <FillTab style={style} set={set} theme={theme} /> : null}
      </div>
    </Dialog>
  )
}
