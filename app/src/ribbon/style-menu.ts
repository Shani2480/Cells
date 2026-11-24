import type { MenuItem } from "@/shell/ContextMenu"
import type { StyleKind } from "@/sheet/custom-styles"
import type { RibbonEnv } from "./registry"

/**
 * A table or PivotTable style's right-click menu in its gallery: Apply (`apply`, absent in a gallery that makes a new
 * table), Modify and Delete for the workbook's own styles, Duplicate, Set As Default.
 */
export function customStyleMenu(
  env: RibbonEnv,
  kind: StyleKind,
  style: string | null,
  close: () => void,
  apply?: (style: string | null) => void,
): MenuItem[] {
  const styles = kind === "pivot" ? env.c.pivotStyles : env.c.tableStyles
  const own = style !== null && !!styles.get(style)
  const dialog = kind === "pivot" ? "pivotStyle" : "tableStyle"
  const item = (label: string, onSelect: () => void, disabled = false): MenuItem => ({
    kind: "item",
    label,
    disabled,
    onSelect: () => {
      close()
      onSelect()
    },
  })
  return [
    ...(apply ? [item("Apply", () => apply(style)), { kind: "separator" } as const] : []),
    item("Modify...", () => env.dialog(`${dialog}:modify:${style}`), !own),
    item("Duplicate...", () => env.dialog(`${dialog}:duplicate:${style}`), style === null),
    item(
      "Delete",
      () =>
        env.c.shell.confirm(`Do you want to delete the style ${style}?`, () => {
          if (style) styles.remove(style)
        }),
      !own,
    ),
    { kind: "separator" },
    item("Set As Default", () => styles.setDefault(style ?? styles.builtInDefault), style === null),
  ]
}
