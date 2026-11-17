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
- [x] Part 4 — Data tools: Advanced Filter, Consolidate, What-If Analysis (Goal Seek, Scenario Manager, Data
      Table), Forecast Sheet, Refresh All and Queries & Connections for text and CSV imports, Edit Links
- [x] Part 5 — Links and text: Link (Ctrl+K, the cell menu's Link ›, following and editing links, Excel files),
      Symbol, Header & Footer (opens Page Setup's tab until Page Layout view exists)
- [x] Part 6 — Views and windows: New Window, Arrange
      All, Hide / Unhide, View Side by Side and synchronous scrolling, Switch Windows
  - [x] Grouped sheets: Shift/Ctrl+click tabs, Select All Sheets / Ungroup Sheets, Ctrl+Shift+PgDn/PgUp, edits
        repeated on each sheet as one undo step, Delete and Hide on the group, [Group] in the title, one-sheet
        commands greyed out, Fill › Across Worksheets, the group saved in Excel files (tabSelected)
  - [x] Split: from the active cell, split bars dragged or double-clicked away, each pane's own wheel and scroll
        bars, Freeze Panes ⇄ Split, saved in Excel files (pane state "split", twips)
  - [x] Page Break Preview: shading, blue page outlines (dashed automatic, solid manual), Page N, dragging breaks
        and the print area's edge, its own zoom, status bar view buttons, view and zooms in Excel files
  - [x] Page Layout view and Ruler: rows and columns on pages as they print, margins, header and footer (click opens
        Header/Footer), blank pages after the data, rulers in in/cm, no frozen panes or split in this view
  - [x] Custom Views: Add (print settings, hidden rows and columns), Show, Delete; read from and written to Excel files
        (customWorkbookViews, customSheetViews, Z_<guid>_.wvu names), checked in Excel 365
  - [x] Windows (decided 2026-11-17: windows are panels inside the Cells window, not OS windows): New Window (a
        second view of the workbook: its own sheet, selection, scroll and zoom; "Book1:2" titles), Arrange All
        (Tiled, Horizontal, Vertical, Cascade), Hide / Unhide, View Side by Side with Synchronous Scrolling and
        Reset Window Position, Switch Windows
- [x] Part 7 — Review: Spelling and Thesaurus with local dictionaries (Hunspell word lists, nothing online),
      Workbook Statistics, Check Accessibility, Allow Edit Ranges
  - [x] Spelling (F7): sheet or selection, Ignore Once/All, Add to Dictionary (shared with Inkspire), Change/Change
        All, Excel's end message; English Hunspell (SCOWL, dictionary-en) bundled. Thesaurus (Shift+F7) pane from
        WordNet 3.1 (scripts/build-thesaurus.mjs, 1.5 MB), meanings, synonyms, antonyms, look up, Insert
  - [x] Check Accessibility: pane with errors and warnings (alt text, table header, contrast, merged cells, default
        sheet names, unclear link text), go to each, recommended actions, Why/How to fix; status bar indicator;
        chart Alt Text (description, Mark as decorative) in xlsx, checked in Excel 365
  - [x] Allow Edit Ranges: New, Modify, Delete, Apply, Protect Sheet; Unlock Range asks a range's password once;
        ranges follow inserted rows and columns; read from and written to Excel files, checked in Excel 365
  - [x] Workbook Statistics (sheet and workbook counts; pictures and other objects join the list with Part 8)
- [x] Part 8 — Objects
  - [x] Text Box, Shapes, Pictures, Icons, WordArt: one object list per sheet (charts included) in stacking order;
        Shape Format and Picture Format tabs; Arrange (order, Align, Group, Rotate, Flip); Selection Pane; Alt Text
        for every object; xlsx read/write keeping Excel's own objects exactly unless edited; printing; statistics
  - [x] Sheet Background (tiled behind the cells, screen only; dimmed in dark mode; xlsx `<picture>`)
  - [x] Themes: Themes, Colors and Fonts galleries, Browse for Themes, writing the theme part; cells, cell styles,
        table styles, charts and shapes follow the theme
- [ ] Part 8b — More objects (decided 2026-11-17: build them), smallest first; Excel-made reference files in
      packages/io/test/fixtures/excel (sigline, oleobject, model3d, smartart)
  - [x] Signature Line: Insert › Text › Signature Line, Signature Setup (signer, title, email, instructions, comments,
        sign date), drawn like Office's; VML in xlsx (`o:signatureline`), Excel's kept
  - [x] Object: Insert › Text › Object › Create from File (as an icon); opens in its own app on double-click;
        xlsx `oleObjects` + VML + embeddings (a file in an OLE package; Office files as they are)
  - [x] Equations: Insert › Equation and its gallery; linear (UnicodeMath) and professional display; the
        Equation tab's structures and symbols; OMML in xlsx text boxes with Excel's text fallback
  - [x] 3D Models: Insert › 3D Models from a file (.glb, .gltf, .fbx, .obj, .3mf, .ply, .stl; other formats turned
        into .glb with three.js, as Office keeps them); drawn with WebGL (three.js, loaded only when a model is drawn),
        Office's camera, lights and turn order (x, then y, then z) matched against Excel's pictures; turned with the
        handle in the middle, 3D Model tab (Change, Reset / Reset and Size, Views, X/Y/Z turns, Alt Text, Arrange,
        Size); xlsx `am3d:model3d` in mc:AlternateContent with its picture fallback, checked in Excel (no repair,
        Shape.Type 30, same rotation). Later: Pan & Zoom (Excel keeps the pan in a part of `am3d` not worked out yet),
        Stock 3D Models (online), animated models, models in groups (written as their picture)
  - [x] SmartArt: Excel's SmartArt shown as Office drew it (its cached drawing) and kept whole while unchanged;
        Insert › SmartArt (Choose a SmartArt Graphic) with Cells' own layout definitions written in Office's diagram
        layout language and checked in Excel (Basic Block List, Basic Process, Basic Cycle, Grid Matrix: the same places
        and sizes in Cells as in Excel), Cells' colour definitions (Colorful, Accent 1–6) and Office's Simple Fill;
        Text Pane (Enter, Tab / Shift+Tab, Backspace, arrows; typing drawn live, one undo step a line); SmartArt
        Design (Add Shape After/Before, Promote, Demote, Move Up/Down, Text Pane, Change Layout, Change Colors, Reset
        Graphic, Convert to Shapes) and Format (Alt Text, Arrange, Size). Written without a cached drawing, so Excel
        lays it out when it opens it. Later: more layouts (pyramid: Excel falls back on Cells' `pyra` definition, so
        it needs more work; Venn: Excel lays out one with a negative spacing only for three circles; hierarchies),
        SmartArt Styles (3D and effects), Excel's text kept as formatted when only some points change (now the data is
        written anew), Office layouts Cells doesn't have drawn as they are (now as the nearest of Cells' own after an
        edit; Excel lays them out properly again)
- [ ] Part 9 — Sparklines, Slicers and Timelines for tables, Quick Analysis, filter and sort by icon
  - [x] Filter and sort by icon: Filter/Sort by Color › by Cell Icon, the cell menu's icon commands, Sort On
        Conditional Formatting Icon; Filter and Sort by Color see conditional formats' colours; iconFilter in xlsx
        (checked against a file Excel 365 made)
  - [x] Sparklines step 1: shown and kept (core sparklines.ts, x14 sparklineGroups in xlsx, drawn in cells; checked
        against Excel's picture of the same file and opened in Excel 365)
  - [x] Sparklines step 2: Insert › Sparklines (Create Sparklines), the Sparkline tab (Edit Data, Type, Show, Style,
        colours, Axis, Group / Ungroup, Clear). Not yet: the cell menu's Sparklines ›, copying and filling cells with
        sparklines, picking ranges on the sheet while the dialog is open
  - [x] Slicers step 1: table slicers from Excel files shown and used (core slicers.ts: items, filter, Excel's 14
        styles from the theme; x15 slicer caches, slicers parts and sle:slicer frames written anew; click, Ctrl+click,
        Shift+click, Multi-Select, Clear Filter, wheel scrolling; follow a renamed header). Checked against Excel's
        pictures of all 14 styles and opened in Excel 365. A file with PivotTable slicers keeps all slicers as read.
  - [x] Slicers step 2: Insert › Slicer and Table Design › Insert Slicer (Insert Slicers dialog), the Slicer tab
        (caption, Slicer Settings, Report Connections greyed, style gallery, columns, button height and width, size,
        arrange; Group and Rotate off for slicers). Opened in Excel 365 with every setting. Not yet: the Format
        Slicer pane, hover colours, copying a sheet links its slicers to the copied table
  - [ ] Quick Analysis (Ctrl+Q)
  - Timelines go with Part 10 (Excel only links them to PivotTables)
- [ ] Part 10 — PivotTables and PivotCharts (the largest: its own plan when we get there)
- [ ] Help tab: Help (a local page of the shortcuts and features), Contact Support and Feedback (a pre-filled
      GitHub issue the person reviews), Show Training
- [ ] Then: Inkspire's side of the link (`docs/inkspire-todo.md`, I1–I9), done in the Inkspire project

Staying out, with the reason in the code: Power Query's online sources, subscriptions and the Data Model
(Relationships, Manage Data Model) need Microsoft's engine and services; co-authoring, Hide Ink and Screenshot need
an account, a pen layer or a desktop capture Cells doesn't have; View Code and Macros need VBA, which Cells doesn't
run (macros in `.xlsm` files are kept). SmartArt, 3D Models, Equations, Signature Line and Object (OLE) are
built in Part 8b.

## Later (from Phase 7)

- Spelling and Thesaurus: more languages as optional downloads (first-run setup / Settings), with the language
  picker in the Spelling dialog; Excel's proofing options (ignore words in UPPERCASE, with numbers, addresses) in
  File › Options › Proofing; AutoCorrect; checking text in charts, text boxes, comments and the header/footer.

- Shapes and pictures: no Crop tool, picture corrections, colour or artistic effects yet; no Shape Styles or WordArt
  Styles galleries on the Format tabs (the colour menus are there); Edit Shape / Edit Points, freeform drawing and
  connectors that stay glued to shapes are missing; adjustment handles (the yellow dots) aren't shown; text in a
  shape is typed as plain text (one format for all of it); Graphics Fill for icons; the Size dialog (Format Shape
  pane). Spelling doesn't check text in shapes yet.

- Themes: Effects (the gallery is shown but not built: Cells' shapes don't use the theme's effect styles yet);
  Customize Colors and Customize Fonts (custom sets); Save Current Theme; Reset to Theme from Template. Conditional
  formats and colours in charts read from a file keep their colours (Excel's follow the theme). .xls files keep their
  colours as read (BIFF8 has no theme references).
- Equations: typing happens in the linear form (Office's Linear view) and builds up when the editor closes; Office
  edits built up in place, with the caret inside the structure. Not built: Ink Equation, Unicode/LaTeX conversion,
  Normal Text, the equation's own font size and alignment buttons, Alt+= (it's AutoSum in Excel anyway). Text beside
  an equation in one paragraph becomes the equation's normal text. Display-size ∫ and ∑ come from the font: bundle a
  math font (STIX Two Math) if Linux/Android ones look small.
- Object: Link to file (an oleLink in an externalLink part, numbered with the workbook links); Create New (needs the
  other app's object server); objects shown as their content rather than an icon (Excel's picture is an EMF, which
  Cells can't draw: such objects show as icons); an Office document's own name isn't kept by Excel, so it comes back
  as "Microsoft_Word_Document.docx". Android can't open embedded files yet (no cache file sharing).
- Signature Line: Sign (a digital ID and the file's XML signature, _xmlsignatures) and Add Signature Services aren't
  built; Excel redraws the line's picture itself when it saves.
- Windows: only the active window's view is saved (Excel writes a workbookView per window and reopens them all);
  the sheet tabs and status bar follow the active window instead of each panel having its own; Cascade's panels
  can't be dragged or resized; Hide is greyed out for the last visible window (Cells shows one workbook at a time).
- Sheet Background: .xls files open without theirs (the BIFF BITMAP record isn't read).

- Check Accessibility: contrast ignores conditional formats, table styles and number-format colours; Excel's
  "Keep accessibility checker running" setting and its Tips aren't there; Alt Text moves to the chart's Format tab
  (and to pictures and shapes) with Part 8.

- Allow Edit Ranges: Permissions... (Windows users and groups) isn't offered; a file's permissions are kept as read.
  Refers to cells is typed (no collapse-to-pick button yet).

- Page Layout: Excel edits the header and footer in place on the page (Cells opens Page Setup › Header/Footer);
  print titles aren't repeated on each page; the ruler setting isn't saved with the sheet (sheetView showRuler).

- Custom Views: a view's filter settings aren't saved or shown (Excel's .wvu.FilterData is kept as read), nor are
  split panes or hidden sheets.

- Links: Create New Document isn't there (Cells shows one workbook per window); Browse for File is on the desktop
  only (Android's picker gives links that don't last). A click follows a link anywhere in its cell (newer Excel
  follows only on the text). The Symbol dialog's Font list stays "(normal text)": a cell has one font, so a symbol
  takes the cell's. Links don't move with a cut and paste or Insert/Delete Cells (they do with whole rows and
  columns), and a pasted cell doesn't bring its link.

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
- Refreshing a linked text file overwrites from the range's top-left and clears what the old data covered beyond the
  new (Excel's "Overwrite existing cells with new data, clear unused cells"). Excel's default, "Insert cells for new
  data, delete unused cells", and filling down formulas beside the data are not built yet. Neither are Power Query
  queries (they are kept in the file as they were read).
- Workbook Links: formulas show a link by its file name ('[Prices.xlsx]Sheet1'!B2); Excel shows the full path while
  the other file is closed. Break Link turns cells into values but leaves defined names that use the link (they show
  #REF! once it's gone); links in conditional formatting and data validation formulas are kept as written. Links
  to text files and Excel's DDE/OLE links aren't built (a file with DDE/OLE links keeps all its links as read).
  Opening the other workbook from the pane isn't built either (Cells shows one workbook per window).

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
