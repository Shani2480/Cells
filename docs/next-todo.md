# Next

## Phase 1 — done

- [x] Scaffold (workspaces, Tauri 2, CI, boundary check, identity)
- [x] Core model, operations, undo/redo with selection restore
- [x] Viewport maths for 1,048,576 × 16,384
- [x] Canvas renderer
- [x] Selection and fill handle
- [x] Editing (type, F2, Enter/Tab, Esc, Alt+Enter, Ctrl+Enter, Delete)
- [x] Excel keyboard navigation
- [x] Clipboard (TSV + HTML, own format, marching ants, tiling)
- [x] Insert/delete, resize, autofit, hide/unhide
- [x] Minimal shell (title bar, formula bar, grid, status bar)
- [x] Excel tool skeleton (`tools/excel-run.py`, kept local)

## Phase 2 — done

- [x] Lexer and parser (A1, R1C1, 3-D, arrays, spill `#`, `@`), printing
- [x] Values and coercion
- [x] Dependency graph, dirty marking, volatiles, circular references, manual mode, F9
- [x] Dynamic arrays and spills
- [x] Structural edits rewrite formulas (insert/delete, move, copy, fill)
- [x] 262 functions, Excel conformance suite (`tests/excel-golden/`)
- [x] Formula bar and in-cell editing: colours, Point mode, F4, autocomplete, tooltips, Ctrl+`, Ctrl+Shift+Enter
- [x] Engine in a Web Worker
- [ ] Chain benchmark under 200 ms on a quiet machine (measured 150–200 ms quiet, more under load)

## Phase 3 — done, waiting for the go-ahead for Phase 4

- [x] Ribbon, KeyTips, Ctrl+F1, Quick Access Toolbar, Search (Alt+Q), backstage with What's New
- [x] Formula bar: Name Box with names, Insert Function, expand and resize
- [x] Sheet tabs
- [x] Cell formatting, number formats (594 Excel TEXT cases), Format Cells
- [x] Conditional formatting with Manage Rules
- [x] Freeze panes, zoom, gridlines/headings, dark mode
- [x] Notes, threaded comments, Protect Sheet, Protect Workbook
- [x] Context menus and shortcuts
- [x] Tablet gestures (grippers, long press, double tap, pinch, keyboard-aware editor)
- [x] Version 0.3.0 with release notes
- [ ] Split panes, Page Break Preview / Page Layout view, New Window
- [ ] Function Arguments dialog; Name Manager, Create from Selection, Apply/Paste Names
- [ ] First-run setup screen (theme, your name for notes and comments, AI opt-in once there is AI). Your rules ask
      for it in every app, but it isn't in PLAN.md yet. Proposed for Phase 5 with Settings.

## Later (from Phase 3)

- Notes and comments don't move with Insert/Delete Cells (shift) or a cut and paste. They do move with whole rows
  and columns.
- "Select locked cells" off isn't enforced: the cursor can still go to locked cells.
- Rows don't refit when a column width changes; edits over 20,000 rows don't refit rows at all.
- Wheel scrolling moves by pixels, not by rows like Excel.
- Icons in the conditional-formatting menu rows; avatar colours for comment authors (now picked from the name).

## Later (from Phase 2)

- Version and What's new page in File → Account (Phase 3 backstage), shown once after an update; the notes are
  in `app/src/about/release-notes.ts`

- LAMBDA and its helpers (MAP, REDUCE, BYROW, …)
- R1C1 display option, defined names (Name Manager), iterative calculation
- Calculation options in the ribbon (Formulas tab, Phase 3)
- ROMAN's concise forms; legacy `{=…}` array formulas
- RATE, IRR and XIRR stopping where Excel's solver does
