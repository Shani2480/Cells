/**
 * What changed in each version, in plain words, newest first. Shown on the What's new page (File → Account, with
 * the backstage in Phase 3) and once after an update. A test checks the newest entry matches the version built.
 */
export type ReleaseNote = { version: string; date: string; notes: string[] }

export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "0.2.0",
    date: "2026-11-14",
    notes: [
      "Formulas calculate: 262 Excel functions, dynamic arrays that spill, circular reference warnings and manual calculation with F9.",
      "References in a formula are coloured, with matching boxes on the grid; point at cells with the mouse or the arrow keys.",
      "While typing a formula, a list suggests functions (Tab inserts one) and a tip shows the function's arguments.",
      "F4 cycles a reference through $A$1, A$1, $A1 and A1. Ctrl+` shows the formulas instead of their results.",
      "Inserting or deleting rows and columns, moving cells and copying formulas update the references, as in Excel.",
      "Calculation runs in the background, so the sheet stays responsive while it works.",
    ],
  },
  {
    version: "0.1.0",
    date: "2026-11-14",
    notes: [
      "The first version: a fast grid of 1,048,576 rows by 16,384 columns with Excel's keyboard shortcuts.",
      "Typing, editing, fill handle, copy and paste, insert and delete, row heights and column widths, hide and unhide.",
      "Undo and redo up to 100 steps.",
    ],
  },
]

export const CURRENT_VERSION = RELEASE_NOTES[0].version
