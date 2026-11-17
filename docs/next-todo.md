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

## Phase 5 — Inkspire integration (done in Cells, 0.5.0; Inkspire's side is `docs/inkspire-todo.md`)

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
- [x] Wrap-up docs (changes.md, verify-list, What's new for 0.5.0)

## Phase 6 — Packaging and phone UI (done, 0.6.0)

Decided 2026-11-16: Linux, Windows, macOS and Android as planned, no iOS; speed measured in the emulator for now.

- [x] Desktop installers from one tag (`.deb`, AppImage, Flatpak, Windows NSIS, macOS `.dmg`; `docs/release.md`),
      file associations, files opened from the file manager or Finder
- [x] macOS menu bar (Cells, File, Edit, View, Format, Window, Help; `src-tauri/src/menu.rs`), Quit that asks to
      save, macOS's own ⌘H/⌘M/⌘Q/⌘,/⌘` left to it, and ⌃H, ⌃G, ⌃` for Replace, Go To and Show Formulas as in Excel
      for Mac. Not yet: the rest of Excel for Mac's own keys (built from memory; see verify-list)
- [x] Android build from `main`: generated project committed (`app.cells`), system bars and keyboard kept clear,
      Cells' icon, `.apk`/`.aab` in the release workflow signed with Inkspire's key once its secrets are copied in
- [x] Android files: Open and Save As through the system's document picker with lasting read and write access
      (`MainActivity.kt`, `app/src/platform/android-files.ts`), real file names, Recent after a restart, "Open with"
      and Share from other apps (at launch and while running), hardware-keyboard shortcuts, no on-screen keyboard
      until editing on a touch screen. Checked on a real Android tablet. Stylus not tested yet.
- [x] Shell choice by screen size and pointer (`app/src/phone/shell-choice.ts`): phone shell under 600 px on the
      smaller side with a finger, else the full window; switches live on resize, rotation and split screen
- [x] PhoneShell (`app/src/phone/`): top bar (File, name, Undo, Redo, Find), 30 px headings, the formula bar with a
      formula key row over the keyboard (`= ( ) , : $ + - * / & "`, fx, Enter), sheet switcher with New Sheet, the
      ribbon's tabs as a command palette in a bottom sheet (tab picker, menus and fields in place, swipe to close),
      fling scrolling; long press, pinch and double tap from the tablet work. Status bar follows the theme.
- [x] Phone dialogs, the File view and setup full screen (`app/src/styles/phone.css`)
- [ ] Simplified phone forms of mouse-only features (chart elements, drag-to-resize panes)
- [x] Emulator smoke test (`scripts/android-smoke.sh`): open, edit, save and read back, frame times, Back; phone and
      tablet screenshots in `docs/images`. It found two bugs, both fixed: the keyboard turned a landscape tablet into
      the phone layout mid-typing, and Back ended Cells (with unsaved work) instead of closing what was open
- [x] The smoke test in CI (`.github/workflows/android-smoke.yml`: phone and tablet emulators, on pushes to main
      that touch the app); not run yet, as the account's Actions are blocked by billing
- [x] Long press on a phone: an action sheet with the context menu's items (`app/src/phone/ActionSheet.tsx`)
- [x] Wrap-up docs (changes.md, verify-list, What's new for 0.6.0)

## Phase 7 — The commands marked "not built" (0.7.0)

Asked 2026-11-16: build what the ribbon and the right-click menus show as "Not built yet." (about 90 ribbon
commands, 10 menu items), then Inkspire's side of the link. In parts, each one shippable, most useful first. Each
follows Excel's names, dialogs and behaviour; what stays out is listed at the end with the reason.

- [x] Part 1 — Menus wired to what exists: the cell menu's Paste Special, Filter › (Clear Filter From, Reapply, by
      the cell's value, colour or font colour) and Sort › (A to Z, Z to A, colour on top, Custom Sort), Pick From
      Drop-down List (and Alt+Down on a plain cell); Sort A to Z / Z to A on the Data tab (a duplicate command had
      replaced them); the Comments pane (Show Comments); the Office Clipboard pane (Home › Clipboard launcher).
      Moved on: by icon (needs icon filters, Part 9) and Select All Sheets (needs grouped sheets, Part 6)
- [x] Part 2 — Formula auditing: Name Manager, Create from Selection, Trace Precedents / Dependents and Remove
      Arrows (drawn on the sheet), Error Checking, Evaluate Formula, Watch Window
- [x] Part 3 — Outline: Group / Ungroup rows and columns, the outline bar with its levels, Show / Hide Detail,
      Subtotals, Settings; read and written in Excel files (today kept but not shown). Also Auto Outline and
      Ctrl+8. Not yet: Settings' Automatic styles / Apply Styles (styling summary rows)
- [ ] Part 4 — Data tools: Advanced Filter, Consolidate, What-If Analysis (Goal Seek, Scenario Manager, Data
      Table), Forecast Sheet, Refresh All and Queries & Connections for text and CSV imports, Edit Links
- [ ] Part 5 — Links and text: Link (Ctrl+K, the cell menu's Link ›, following and editing links, Excel files),
      Symbol, Header & Footer (opens Page Setup's tab until Page Layout view exists)
- [ ] Part 6 — Views and windows: grouped sheets (Select All Sheets, Shift/Ctrl+click tabs), Page Break Preview, Page Layout view, Custom Views, Ruler, New Window, Arrange
      All, Split, Hide / Unhide, View Side by Side and synchronous scrolling, Switch Windows
- [ ] Part 7 — Review: Spelling and Thesaurus with local dictionaries (Hunspell word lists, nothing online),
      Workbook Statistics, Check Accessibility, Allow Edit Ranges
- [ ] Part 8 — Objects: Text Box, Shapes, Pictures, Icons, WordArt; Bring Forward / Send Backward, Align, Group,
      Rotate, Selection Pane; Sheet Background; Themes, theme colours, fonts and effects
- [ ] Part 9 — Sparklines, Slicers and Timelines for tables, Quick Analysis, filter and sort by icon
- [ ] Part 10 — PivotTables and PivotCharts (the largest: its own plan when we get there)
- [ ] Help tab: Help (a local page of the shortcuts and features), Contact Support and Feedback (a pre-filled
      GitHub issue the person reviews), Show Training
- [ ] Then: Inkspire's side of the link (`docs/inkspire-todo.md`, I1–I9), done in the Inkspire project

Staying out, with the reason in the code: Power Query's online sources, subscriptions and the Data Model
(Relationships, Manage Data Model) need Microsoft's engine and services; co-authoring, Hide Ink and Screenshot need
an account, a pen layer or a desktop capture Cells doesn't have; View Code and Macros need VBA, which Cells doesn't
run (macros in `.xlsm` files are kept). SmartArt, 3D Models, Equations, Signature Line and Object (OLE) are
decided when Part 8 is done.

## Later (from Phase 7)

- The Name Box doesn't take a list of ranges (B1,B3) yet.
- Data tables in .xls files open as their values (the BIFF TABLE record isn't read).
- An Advanced Filter in place isn't saved as one (its hidden rows are): after reopening, Clear and Reapply don't
  know it (Excel keeps it in the _FilterDatabase and Criteria names).
- Consolidate works with ranges in this workbook only (Excel's Browse... for other workbooks isn't there); settings
  that name another workbook are kept in the file as they were read.
- FORECAST.ETS and its family give results close to Excel's, not equal: the season length and perfect patterns match
  Excel, but on noisy data the fitted smoothing factors differ (forecasts within about 0.5%, confidence widths
  within about 10%, from a check of 8 patterns against Excel 16). Excel doesn't publish its fitting.
- Forecast Sheet's chart is a plain line or column chart; Excel draws the confidence bounds thinner.

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
