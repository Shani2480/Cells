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

## Phase 4 — done (0.4.0)

- [x] xlsx open and save (unknown parts kept), big files with progress, CSV/TXT import and export, File views,
      recent files, AutoRecover
- [x] Sort and Filter (AutoFilter, custom sort, filter by colour)
- [x] Find and Replace, Go To Special, Remove Duplicates, Text to Columns
- [x] Data Validation
- [x] Tables (Format as Table) with structured references
- [x] Flash Fill
- [x] Charts (Insert galleries, Chart Design tab, Select Data, xlsx read and write)
- [x] Printing and PDF: page setup on every sheet (saved in xlsx and read back by Excel), the Page Layout tab's Page
      Setup / Scale to Fit / Sheet Options, the Page Setup dialog (four tabs, custom header and footer), print area
      and titles, page breaks (dashed on the sheet), File › Print with a preview, File › Export to PDF.
- [x] Wrap-up docs (changes.md, verify-list, What's new for 0.4.0)

## Phase 5 — Inkspire integration (in progress, version 0.5.0 not shipped yet)

Decided 2026-11-15: Cells gets AI features using Inkspire's providers and keys; regional formats are applied in
Cells (typing and showing dates and numbers); each app keeps its own recent-files list.

- [x] Shared settings file (`docs/shared-settings.md`), File › Options (General, Advanced), first-run setup, notes
      and comments signed with the user name, Office Theme
- [x] Regional formats applied: typed numbers and dates, number formats, General, the short date and the formula bar
- [x] `cells://` links and a single instance (open a file at a range, from Inkspire; files from the command line)
- [x] Copy a range (formatted HTML with its source) or a chart (Inkspire's chart markup and a picture) for Inkspire;
      a chart pastes back into Cells as a copy
- [x] Linked ranges: Cells writes what Inkspire's linked ranges hold as they change; "Edit Data" opens a chart's
      workbook in Cells and saves it back (`docs/linked-ranges.md`)
- [x] AI: formula from a description, explain a formula, suggest a chart or summary (Inkspire's providers and keys;
      opt-in in the setup and Options; `docs/ai.md`). Not yet: Inkspire's built-in engine in Cells
- [x] Shared code packages for Inkspire to adopt: `@cells/ooxml` (new), `@cells/chart`, and the formula engine's
      `calculateGrid`/`calculateFormula` in `@cells/core` (`docs/shared-packages.md`)

## Later (from Phase 4)

- Page Layout view and Page Break Preview (dragging page breaks).
- PDF pages are pictures of the pages, so their text can't be selected or searched.
- Even-page and first-page headers from Excel files are kept but can't be edited.
- Black and white printing is saved but not applied to the drawing.
- Charts: Format tab, Insert Chart dialog, Move Chart, chart sheets.
- Flash Fill's live suggestion while typing.
- Sparklines and `.ods` (both marked optional in PLAN.md).

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
