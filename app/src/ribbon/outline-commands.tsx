import type { Command, RibbonEnv } from "./registry"

/** Group or Ungroup the selection: whole rows or columns go at once, anything else asks which (Excel's dialog). */
function groupOrAsk(env: RibbonEnv, ungroup: boolean) {
  const axis = env.c.outline.axisOfSelection()
  if (!axis) return env.dialog(ungroup ? "outlineUngroup" : "outlineGroup")
  if (ungroup) env.c.outline.ungroup(axis)
  else env.c.outline.group(axis)
}

/** Data › Outline. */
export function outlineCommands(): Command[] {
  return [
    {
      id: "OutlineGroupMenu",
      label: "Group",
      description: "Tie a range of cells together so that they can be collapsed or expanded.",
      shortcut: "Alt+Shift+Right",
      run: (env) => groupOrAsk(env, false),
      menu: ["OutlineGroup", "OutlineAutoOutline"],
    },
    { id: "OutlineGroup", label: "Group...", run: (env) => groupOrAsk(env, false) },
    {
      id: "OutlineAutoOutline",
      label: "Auto Outline",
      description: "Group the rows and columns that formulas add up.",
      run: (env) => env.c.outline.autoOutline(),
    },
    {
      id: "OutlineUngroupMenu",
      label: "Ungroup",
      description: "Ungroup a range of cells that were previously grouped.",
      shortcut: "Alt+Shift+Left",
      run: (env) => groupOrAsk(env, true),
      menu: ["OutlineUngroup", "OutlineClear"],
    },
    { id: "OutlineUngroup", label: "Ungroup...", run: (env) => groupOrAsk(env, true) },
    { id: "OutlineClear", label: "Clear Outline", run: (env) => env.c.outline.clearOutline() },
    {
      id: "OutlineSubtotals",
      label: "Subtotal",
      description: "Total several rows of related data together by automatically inserting subtotals and totals.",
      run: (env) => env.dialog("subtotal"),
    },
    {
      id: "OutlineShowDetail",
      label: "Show Detail",
      description: "Expand a collapsed group of cells.",
      run: (env) => env.c.outline.detail(true),
    },
    {
      id: "OutlineHideDetail",
      label: "Hide Detail",
      description: "Collapse a group of cells.",
      run: (env) => env.c.outline.detail(false),
    },
    {
      id: "OutlineSettings",
      label: "Outline Settings",
      description: "Show the Outline dialog box.",
      run: (env) => env.dialog("outlineSettings"),
    },
  ]
}
