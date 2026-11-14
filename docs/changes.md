# Changes

Newest first. Each entry: what changed, how it was checked, what is still open.

## 2026-11-14 — Phase 1: grid, editing, clipboard, rows and columns

**Built**

- npm workspaces: `packages/core` (pure TypeScript: no DOM, Node, React or Tauri, enforced by ESLint and
  `scripts/check-boundaries.mjs`), `packages/platform` (interfaces), `app/` (React shell), `src-tauri/` (Tauri 2).
- Core model: sparse column/chunk cell store, interned styles, Fenwick-tree row/column axes, serialisable operations
  that return their inverse, a workbook with grouped undo/redo (100 levels, as Excel) that restores the selection.
- Excel behaviour, checked in Excel (16.0, build 17932): Ctrl+arrow data edges (also across hidden rows), Current
  Region (grows through diagonals), AutoFill series (numbers by trend, day/month lists, numbered text, quarters,
  ordinals, Ctrl to invert), 15 significant digits on entry (cut off, not rounded).
- Canvas grid: headers that highlight with the selection, text overflow into empty neighbours, `####` for numbers
  that don't fit, selection with the active cell cut out, fill handle and fill preview, marching ants, resize
  guides with Excel's width/height tooltip, auto-fit on double-click.
- Editing: type-to-replace, F2 (Enter/Edit modes), Enter/Tab with Excel's Tab-return, Esc, Alt+Enter,
  Ctrl+Enter, Delete, Ctrl+D/R, fill handle (drag, double-click to fill down to the data).
- Keyboard: arrows, Ctrl+arrows, Home/Ctrl+Home/Ctrl+End, PgUp/PgDn (Alt for horizontal), End mode, Ctrl+A (region
  then sheet), Ctrl+Space/Shift+Space, Ctrl+*, F5/Ctrl+G, Ctrl+-/Ctrl++, Ctrl+9/0 and Ctrl+Shift+9/0, Ctrl+Z/Y.
- Clipboard: TSV + HTML out, own copies paste with full content, cut pastes as a move, tiling when the target is a
  multiple of the copy, outside text parsed as typed input.
- Rows and columns: insert/delete (Insert/Delete dialogs with the shift choices), height/width dialogs in Excel's
  units (points, characters), hide/unhide.
- Shell: title bar (logo, Undo/Redo with their labels), formula bar (Name Box with "3R x 2C" while dragging,
  Cancel/Enter, fx), status bar (Ready/Enter/Edit, End Mode, Average/Count/Sum), context menu, Go To dialog.
- App icon: `design/app-icon.svg` (master) → `src-tauri/icons`, favicon and the title-bar mark.
- CI: lint (ESLint, Prettier, boundaries), typecheck, Vitest, build, cargo test.

**Deviation from PLAN.md**

- The engine runs on the main thread in Phase 1, behind the same operation interface the Worker will use. There is
  no calculation yet, so a Worker would only add copying; it moves into `engine-host` with the calc engine in
  Phase 2.

**Measured** (100,000 × 20 sheet of made-up data)

- Core benchmarks: Ctrl+Down 4 ms, insert a row at the top 56 ms (undo 49 ms), whole-column stats 6.5 ms, copy
  10k × 20 as TSV 102 ms, 1,000 undo steps 3.5 ms.
- In WebKitGTK (MiniBrowser on Xvfb, software rendering, dev build): data on screen 0.6 s after the empty app
  (about 0.9 s from navigation); one draw 5 ms median, 6 ms p95; scrolling frame p95 18–31 ms (vsync-bound
  at 17 ms in most runs; the spread comes from the software compositor); Ctrl+Down to row 100,000 painted in
  19–38 ms.

**Open**: see `docs/verify-list.md`.
