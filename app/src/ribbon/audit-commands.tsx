import type { Command } from "./registry"

/** Formulas › Defined Names (Name Manager, Create from Selection) and Formula Auditing. */
export function auditCommands(): Command[] {
  return [
    {
      id: "NameManager",
      label: "Name Manager",
      description: "Create, edit, delete and find all the names used in the workbook.",
      shortcut: "Ctrl+F3",
      run: (env) => env.dialog("nameManager"),
    },
    {
      id: "NameCreateFromSelection",
      label: "Create from Selection",
      description: "Automatically generate names from the selected cells.",
      shortcut: "Ctrl+Shift+F3",
      run: (env) => env.dialog("createNames"),
    },
    {
      id: "TracePrecedents",
      label: "Trace Precedents",
      description: "Show arrows that indicate which cells affect the value of the currently selected cell.",
      run: (env) => {
        const c = env.c
        const input = c.sheet.cells.input(c.selection.active.row, c.selection.active.col)
        if (!c.audit.tracePrecedents() && !input?.startsWith("="))
          c.shell.alert("The active cell must contain a formula that refers to other cells.")
      },
    },
    {
      id: "TraceDependents",
      label: "Trace Dependents",
      description: "Show arrows that indicate which cells are affected by the value of the currently selected cell.",
      run: (env) => {
        if (!env.c.audit.traceDependents() && !env.c.audit.hasArrows)
          env.c.shell.alert("The Trace Dependents command found no formulas that refer to the active cell.")
      },
    },
    {
      id: "TraceRemoveArrowsMenu",
      label: "Remove Arrows",
      description: "Remove the arrows drawn by Trace Precedents or Trace Dependents.",
      run: (env) => env.c.audit.removeArrows("all"),
      menu: ["TraceRemoveArrows", "TraceRemovePrecedentArrows", "TraceRemoveDependentArrows"],
    },
    { id: "TraceRemoveArrows", label: "Remove Arrows", run: (env) => env.c.audit.removeArrows("all") },
    {
      id: "TraceRemovePrecedentArrows",
      label: "Remove Precedent Arrows",
      run: (env) => env.c.audit.removeArrows("precedents"),
    },
    {
      id: "TraceRemoveDependentArrows",
      label: "Remove Dependent Arrows",
      run: (env) => env.c.audit.removeArrows("dependents"),
    },
    {
      id: "ErrorCheckingMenu",
      label: "Error Checking",
      description: "Check for common errors that occur in formulas.",
      run: (env) => env.dialog("errorChecking"),
      menu: ["ErrorChecking", "TraceError"],
    },
    { id: "ErrorChecking", label: "Error Checking...", run: (env) => env.dialog("errorChecking") },
    {
      id: "TraceError",
      label: "Trace Error",
      description: "Show arrows from the cells that cause the error in the active cell.",
      run: (env) => {
        if (!env.c.audit.traceError()) env.c.shell.alert("The active cell doesn't show an error.")
      },
    },
    {
      id: "FormulaEvaluate",
      label: "Evaluate Formula",
      description: "Debug a complex formula, evaluating each part individually.",
      run: (env) => env.dialog("evaluateFormula"),
    },
    {
      id: "WatchWindow",
      label: "Watch Window",
      description: "Monitor the values of certain cells as changes are made to the sheet.",
      run: (env) => env.c.audit.toggleWatchWindow(),
      pressed: (env) => env.c.audit.watchOpen,
    },
  ]
}
