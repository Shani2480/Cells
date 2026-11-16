import { ChevronRight } from "lucide-react"
import type { MenuItem } from "@/shell/ContextMenu"
import { BottomSheet } from "./BottomSheet"

/**
 * A long press on a phone: the right-click menu's own items (App's menuItems) as a bottom sheet, after Excel for
 * phones. Cut, Copy and Paste become big buttons in a row; the rest are finger-sized rows, grouped as the menu's
 * separators group them. Items Cells hasn't built are left out rather than shown dead, as in the command palette.
 */

type Action = Extract<MenuItem, { kind: "item" }>

/** The quick row's labels, in order. */
const QUICK = ["Cut", "Copy", "Paste"]

/** The menu's items as the sheet shows them: the quick row, then groups without empty ones or unbuilt items. */
export function actionSheetItems(items: MenuItem[]): { quick: Action[]; groups: Action[][] } {
  const quick: Action[] = []
  const groups: Action[][] = [[]]
  for (const item of items) {
    if (item.kind === "separator") {
      if (groups[groups.length - 1].length) groups.push([])
      continue
    }
    if (item.disabled) continue
    if (QUICK.includes(item.label) && item.icon) quick.push(item)
    else groups[groups.length - 1].push(item)
  }
  return { quick, groups: groups.filter((group) => group.length) }
}

export function ActionSheet({ label, items, onClose }: { label: string; items: MenuItem[]; onClose: () => void }) {
  const { quick, groups } = actionSheetItems(items)
  const choose = (item: Action) => {
    onClose()
    item.onSelect()
  }
  return (
    <BottomSheet label={label} onClose={onClose} header={<span className="phone-sheet-title">{label}</span>}>
      {quick.length ? (
        <div className="phone-quick-row" role="group" aria-label="Clipboard">
          {quick.map((item) => (
            <button key={item.label} type="button" className="phone-quick" onClick={() => choose(item)}>
              <span className="phone-quick-icon" aria-hidden>
                {item.icon}
              </span>
              <span>{item.label}</span>
            </button>
          ))}
        </div>
      ) : null}
      {groups.map((group, index) => (
        <div key={index} className="phone-group" role="group">
          {group.map((item) => (
            <button key={item.label} type="button" className="phone-row" onClick={() => choose(item)}>
              <span className="phone-row-icon" aria-hidden>
                {item.icon}
              </span>
              <span className="phone-row-label">{item.label.replace(/\.\.\.$/, "…")}</span>
              {item.more ? <ChevronRight size={18} aria-hidden /> : null}
            </button>
          ))}
        </div>
      ))}
    </BottomSheet>
  )
}
