# Changes

Newest first. Each entry: what changed, how it was checked, what is still open.

## 2026-11-16 — Phase 6: packaging, Android tablets and phones (0.6.0)

**Built**

- Installers from one tag (`docs/release.md`, `.github/workflows/release.yml`): `.deb` and AppImage (Ubuntu 22.04),
  a Flatpak made from the `.deb`, a Windows setup `.exe`, a universal macOS `.dmg`, and an Android `.apk` and `.aab`.
  The release text is the version's notes from What's New; the job checks the tag, version and Flatpak metainfo agree.
- File associations: workbooks (`.xlsx`, `.xlsm`, templates) and, as an alternative, `.csv` and `.tsv`. Files from the
  file manager and the Finder open in the running Cells; the Flatpak shares Inkspire's settings folders.
- The macOS menu bar (`src-tauri/src/menu.rs`) with Excel for Mac's ⌘ shortcuts; ⌘H, ⌘M, ⌘Q, ⌘, and ⌘` left to
macOS, ⌃H, ⌃G and ⌃` for Replace, Go To and Show Formulas.
- Android from `main` (`src-tauri/gen/android`, id `app.cells`): system bars and keyboard kept clear and coloured
  with the theme, Cells' icon, signing with Inkspire's key once its secrets are copied in. Files through Android's
  document picker with lasting access (Save writes back, Recent reopens), real names, "Open with" and Share, hardware
  keyboard shortcuts, the on-screen keyboard only once editing starts, and Back that closes what's open or sends
  Cells to the background.
- The shell by screen, not OS (`app/src/phone/shell-choice.ts`): under 600 px on the smaller side with a finger, the
  phone shell; else the full window. It switches live and ignores the on-screen keyboard.
- The phone shell (`app/src/phone/`): top bar, 30 px headings, the formula bar with a row of formula keys over the
  keyboard, the sheet switcher, and the ribbon's tabs as a command palette in a bottom sheet (tab picker, menus and
  fields in place, swipe to close). Dialogs, the File view and setup are full screen. Fling scrolling on touch.
- `scripts/android-smoke.sh` and `scripts/android-frames.mjs`: open, edit, save and read back, frame times inside the
  web view, Back, and the screenshots in `docs/images`.

**Checked**

- Linux packages built here: `.deb`, AppImage and Flatpak install and start; desktop-file and AppStream validators
  pass; on a private Xvfb display a second launch with a path, a `file://` URL and a `cells://` link reach the
  running Cells.
- The macOS menu built once on Linux (every accelerator parses, the menus show), then kept to macOS.
- On a real Android tablet: open from Downloads through the picker, edit and Ctrl+S saved into the same file
  (read back), Recent after a force-stop, "Open with" at launch and while running, double tap shows the keyboard,
  Back, the full ribbon, and 60 fps scrolling (0.7% of frames late).
- In emulators (Android 15, a Pixel 7 phone and a Pixel Tablet): the smoke test passes on both; the phone shell in
  light and dark, portrait and landscape, the palette, formula keys, sheets, full-screen dialogs and the File view.
  The phone emulator scrolls at 60 fps with the host GPU; software rendering drops frames (`docs/release.md`).
- The smoke test found two bugs, fixed before the commit: the keyboard made a landscape tablet switch to the phone
  layout, and Back ended Cells with unsaved work.

**Deviations from PLAN.md**

- The palette's tabs are Excel for phones' and the ribbon's (Home, Insert, Page Layout, Formulas, Data, Review, View),
  not the plan's "Format, Insert, Data, Formulas, View"; the bottom bar's five tools open it on those.
- The emulator smoke test has a CI workflow (`android-smoke.yml`) that hasn't run yet: the account's Actions are
  blocked by billing.
- The real phone measured is a high-end one rather than mid-range: 120 fps scrolling and typing,
  open, edit, Ctrl+S and read back, formula keys over its keyboard, Back. It found the palette's header ignoring a
  drag that began on the tab picker, fixed and rechecked there.
- A long press on a phone opens an action sheet with the context menu's items (Cut, Copy, Paste as big buttons);
  other mouse-first features have no phone form of their own.

**Open**: see `docs/verify-list.md`.

## 2026-11-16 — Phase 5: Inkspire integration (0.5.0)

**Built**

- Shared settings: a file both apps read and write (`docs/shared-settings.md`) with the user name and initials, the
  Office Theme, the region, short date and separators, and the AI choice. File › Options (General, Advanced, AI) and
  a first-run setup that starts from the system's settings or Inkspire's. Notes and comments are signed with the name.
- Regional formats applied: typed numbers and dates, number formats, General, the short date (format 14) and the
  formula bar follow the person's region; formulas and files stay in Excel's en-US form.
- `cells://open` and `cells://edit` links and a single Cells (`docs/app-links.md`): a second launch, a file from the
  file manager or a link from Inkspire goes to the running Cells, which can select a range.
- Copy for Inkspire (`docs/inkspire-clipboard.md`): a range copies as a formatted HTML table carrying its source; a
  chart copies as Inkspire's chart markup plus a picture, and pastes back into Cells as a copy.
- Linked ranges and Edit Data (`docs/linked-ranges.md`): Cells writes what the ranges Inkspire links to hold, a moment
  after they change; Edit Data opens a chart's embedded workbook in Cells ("Chart in Report.docx") and saves every
  change back.
- AI tools (`docs/ai.md`), off until turned on: a formula from words, Explain Formula, Suggest a Chart and Summarize
  Data, with Inkspire's providers (Ollama, LM Studio, OpenAI-compatible, OpenAI, Anthropic) and its keyring entries.
  Every dialog says what it sends where; nothing goes into the workbook until Insert.
- Shared packages (`docs/shared-packages.md`): `@cells/ooxml` (new: zip, XML and OPC for workbooks and Word
  documents), `@cells/chart`, and `calculateGrid`/`calculateFormula` in `@cells/core` for Inkspire's Spreadsheet block
  and formula fields. They build with ES2020 libraries and compile under Inkspire's `tsconfig.json`.

**Checked**

- The real Tauri build on a private Xvfb display: a `cells://` link at startup opens the workbook at the range, the
  linked-range snapshot holds the right values, and an Edit Data link opens "Chart in Report.docx - Cells". A second
  launch or a later link reaching the running Cells wasn't tried there; only reading the launch arguments is tested.
- The AI streaming path against a fake local endpoint (a reply split mid-character across network chunks), the four
  tools through the app with canned replies, and each dialog in light and dark.
- Regional entry and display in tests for US, German, British and ISO settings; the shared settings file's merge
  (only changed keys, maps key by key) keeping a change Inkspire made in between.

**Deviations from PLAN.md**

- Ribbon components and tokens aren't shared packages: the two ribbons have no code in common and only 19 token names
  match (reasons in `docs/shared-packages.md`).
- Inkspire's built-in AI model doesn't run in Cells yet; Cells points to Ollama or LM Studio instead (I9).
- The plan's cross-app checks (a linked table updating in Inkspire, Word opening Inkspire's linked files) wait for
  Inkspire's side, which is listed in `docs/inkspire-todo.md` (I1–I9) and done in the Inkspire project.

**Open**: see `docs/verify-list.md`.

## 2026-11-15 — Phase 4: files and data tools (0.4.0)

**Built**

- Excel files: `.xlsx`, `.xlsm` and templates read and written in `packages/io` (values, shared strings, formulas of
  every kind, styles, sizes, merges, panes, names, conditional formatting, validation, notes and comments, tables,
  filters, print settings, charts). Parts Cells doesn't model (pictures, macros, pivot caches, outline) are written
  back as they came. SHA-512 sheet and workbook passwords from Excel files work.
- File views: Home, New, Open, Save As, Info, Recent with pinning, Export to PDF; Ctrl+N/O/S/W and F12, the save
  prompt before closing or replacing, the workbook's name in the title bar, AutoRecover every two minutes and
  Document Recovery after a crash, progress while opening, plain-word errors and a list of what was left out.
- CSV and text: `.csv` opens at once (UTF-8 detection, semicolon CSVs) and saves back in its own delimiter and
  encoding; `.txt` through the Text Import Wizard; Data › From Text/CSV; Save As CSV UTF-8 and Text, byte for byte as
  Excel 365 writes them. Drag and drop to open.
- Sort (A to Z, Z to A, the Sort dialog with levels, colours, custom lists, case, left to right) and AutoFilter
  (value list with search and date tree, Text/Number/Date Filters, Top 10, dynamic and colour filters, record count).
- Find and Replace, Go To Special, Remove Duplicates, Text to Columns.
- Data Validation with the dialog, alerts, input messages, in-cell lists and Circle Invalid Data.
- Tables: Format as Table with Excel's 60 styles, the Table Design tab, total row, structured references in the
  engine, growing tables and calculated columns.
- Flash Fill (Ctrl+E).
- Charts in `packages/chart` (shared with Inkspire): Insert galleries and Recommended Charts, moving and sizing, the
  Chart Design tab, Select Data, live data, and DrawingML chart parts read and written.
- Printing: page setup on every sheet, the Page Layout tab, the Page Setup dialog with custom headers and footers,
  print area and titles, page breaks shown on the sheet, File › Print with a preview, and PDF export.

**Checked against Excel** (16.0, US locale)

- Every file Cells writes passes the Open XML SDK validator (3.5.1) and opens in Excel with no repair prompt, with
  the same values and formulas. The Excel-made fixtures in `packages/io/test/fixtures/excel` (budget, formats,
  conditional, layout, protected, passwords, filters, validation, features, print) open, save and reopen with the same content, and their theme and document
  properties come back byte for byte.
- CSV and text output byte for byte against Excel's own Save As.
- Sort order, filter results, Find (62 cases) and Replace (7), Remove Duplicates' matching, Text to Columns'
  splitting and which entries each validation rule accepts, against Excel's results.
- Tables (styles, structured references, a filtered column), charts (=SERIES formulas, an untouched Excel chart
  written back as it came) and every print setting, read back by Excel after Cells saved them; page counts match
  Excel's on the printing fixture.

**Deviations from PLAN.md**

- Not built in this phase: Page Layout view and Page Break Preview, sparklines, `.ods`, the chart Format tab, the
  Insert Chart dialog, Move Chart and chart sheets, Flash Fill's live suggestion while typing.
- PDF pages are pictures of the pages, so their text can't be selected.
- The Android system picker is left for Phase 6; the browser build uses the File System Access API.

**Measured**

- 100,000 rows × 10 columns (1M cells of numbers, text, formulas and logicals; 6.4 MB): save 2.0 s, open 2.2 s in
  Node (`packages/io/test/xlsx.bench.ts`, budget 15 s each). In the app the reading runs in the worker with progress.
- Sheet passwords with Excel's 100,000 SHA-512 spins check in under 1.5 s.

**Open**: see `docs/verify-list.md`.

## 2026-11-14 — Phase 3: desktop and tablet UI, formatting (0.3.0)

**Built**

- Ribbon with Excel's tabs, groups and control ids (Home, Insert, Page Layout, Formulas, Data, Review, View, Help),
  groups that collapse as the window narrows, KeyTips, Ctrl+F1, right-click menu (Add to Quick Access Toolbar,
  Collapse the Ribbon). Everything not built is greyed out with "Not built yet." (Insert, Page Layout and Data are
  mostly that until Phases 4–5.)
- Quick Access Toolbar (Undo/Redo by default, Excel's Customize menu, right-click to add or remove), and Search
  (Alt+Q), which runs a command or opens its menu on its tab.
- File backstage with the Account page and What's New. It opens by itself once after an update.
- Formula bar: Name Box with the names list, naming the selection by typing a name, going to names and references;
  fx and Shift+F3 open Insert Function (search, categories, Most Recently Used); expand/collapse (Ctrl+Shift+U) and a
  drag handle to resize.
- Formulas tab: the function library galleries, Define Name, Use in Formula, calculation options, Calculate Now and
  Calculate Sheet, Show Formulas. Defined names are kept through inserting, deleting, renaming and removing sheets.
- Sheet tabs: add, rename, move/copy (drag, Ctrl+drag, dialog), colour, hide/unhide, scroll arrows, Ctrl+PgUp/PgDn,
  right-click menus.
- Formatting: font, size, bold/italic/underline/strike, colours, fill, alignment, indent, wrap, rotation, merge
  variants, borders, cell styles, Format Painter, clear formats. Row heights follow wrapped text, line breaks and
  bigger fonts (rows sized by hand keep their height).
- Number formats: Excel's format language end to end, the Number group, Increase/Decrease Decimal, Format Cells
  (Number, Alignment, Font, Border, Fill, Protection), entry auto-detection.
- Conditional formatting: highlight rules, top/bottom, data bars, colour scales, icon sets, formula rules, and Manage
  Rules with priority and Stop If True.
- View: freeze panes (at the selection, top row, first column), zoom (dialog, 100%, to selection, status bar
  slider, Ctrl+wheel, pinch), gridlines and headings, formula bar toggle, dark mode.
- Review: notes (yellow boxes, Show/Hide, Show All, Previous/Next, Convert to Comments), threaded comments (card with
  replies, edit, delete, resolve), Protect Sheet (Excel's 15 options and password with confirmation), Protect
  Workbook (structure). Locked cells refuse edits with Excel's message.
- Context menus for cells, row and column headers, tabs and the ribbon; the Phase 3 shortcut set.
- Tablet: `pointer: coarse` sizing, tap to select, round grippers to stretch the selection, long-press menu, double
  tap to edit (the keyboard only shows while editing), pinch zoom around the fingers, and the editor scrolled clear
  of the on-screen keyboard.

**Checked against Excel** (16.0, build 17932, US locale)

- Number formats: `packages/core/test/fixtures/excel-text.json` holds 594 TEXT results that Excel gave (66 format
  codes × 8 numbers and a text). All match. Bugs fixed on the way:
  - quoted literals before digits came out reversed;
  - `@` showed nothing for numbers;
  - seconds rounded down (13:37:37 for 13:37:38);
  - fraction padding, `?/100`-style denominators and engineering zero (`000.0E+0`);
  - the minus sign in conditional sections;
  - out-of-range dates now give `#VALUE!` in TEXT.
- Password hash (Excel's legacy hash) and its messages, Protect Sheet defaults, name and Name Box messages.

**Deviations from PLAN.md**

- Not built in this phase:
  - split panes, Page Break Preview, Page Layout view, and New Window (multiple windows on one workbook);
  - the Function Arguments dialog: Insert Function puts `=NAME()` in the cell with the caret inside;
  - Name Manager, Create from Selection, Apply Names and Paste Names. Names are workbook-wide only.
- The Windows option of Protect Workbook is greyed out, as in Excel 365.

**Measured**

- Formatting A1:A100000 through the app's controller (format + undo, best of 3): bold 61 ms, currency 61 ms, fill
  57 ms, font size 57 ms, borders 55 ms; bold on the whole column A also under the 300 ms budget
  (`app/src/sheet/format-budget.test.ts`).
- One frozen-pane frame of a 100,000 × 20 sheet at 1920 × 1080: layout and drawing code 2.5 ms
  (`app/src/grid/draw.test.ts`; the canvas is a recorder there, so the GPU's share is not included).

**Open**: see `docs/verify-list.md`.

## 2026-11-14 — Phase 2: formula and calculation engine

**Built**

- Formula language: lexer and parser for Excel's grammar (A1 and R1C1 references, mixed `$`, ranges, whole rows and
  columns, 3-D and cross-sheet references, `#REF!` references, array constants, the spill `#` and implicit `@`),
  printing back to text, number literals cut to 15 significant digits as Excel does.
- Values and coercion as Excel: text ↔ number ↔ date, logicals, the error values and how each propagates.
- Calculation: dependency graph with per-cell and per-range links, dirty marking and evaluation in dependency order
  without recursion, volatile functions, circular references (the status bar names the cell), automatic and manual
  mode, F9 and Ctrl+Alt+F9, dynamic arrays with spilling, `#SPILL!` and the spill border.
- 262 functions (maths, statistics, logical, information, text, dates and times, lookup including XLOOKUP, XMATCH
  and FILTER/SORT/UNIQUE, financial, aggregate, arrays such as SEQUENCE and MMULT, and LET). LAMBDA is not built.
- Structural edits rewrite formulas: inserting and deleting rows, columns and cells, cut and paste (`moveCells`), copy
  and fill adjust relative references; undo restores the original text.
- The engine runs in a Web Worker (`app/src/engine-host`): the UI applies operations at once and gets the values
  back as patches.
- Formula editing: references coloured in the cell and the bar with matching boxes on the grid, Point mode (arrows,
  click and drag, Shift to extend), F4, the function list while typing (Tab inserts), the argument tooltip with the
  current argument in bold, Ctrl+`(show formulas), Ctrl+Shift+Enter, missing`)` added on entry, the spilled
  cells' formula shown grey in the bar.

**Checked against Excel** (16.0, build 17932, US locale)

- 719 formulas evaluated by Excel in `tests/excel-golden/` (two rejected by Excel and skipped). All match to 15
  significant digits except four listed in `known-differences.json`: RATE (×2), IRR and XIRR, where Excel's solver
  stops earlier than ours. A known difference that starts matching fails the test, so the list stays honest.
- Numeric details found on the way and copied: whole powers by repeated squaring (FV, PV, EFFECT too; PMT, IPMT and
  PPMT use `exp(n·log1p(r))`), FACT multiplying downwards, odd roots of negative numbers, variance as Σx² − (Σx)²/n,
  MOD by `fmod` with the divisor's sign, INT/FLOOR/CEILING rounding to 15 digits first, COUNTIF's numeric criterion
  matching numeric text, AND/OR skipping text arguments, scaling commas after the last digit in TEXT formats, LEFT,
  RIGHT and MID not splitting surrogate pairs.

**Deviations from PLAN.md**

- The UI still holds the whole workbook (without calculating it) next to the worker's copy, rather than only what
  is on screen. It keeps editing and drawing synchronous; memory is the cost. To revisit when files get large
  (Phase 4).
- Ctrl+Shift+Enter enters the formula like Enter: with dynamic arrays Excel 365 spills either way. Legacy array
  formulas (`{=…}`) are not made.
- ROMAN writes the classic form only; the concise forms (ROMAN(x, 1…4)) are not done.
- R1C1 is parsed but there is no R1C1 display option yet; defined names and iterative calculation are not built.
  The calculation-mode switch has no UI until the ribbon (Phase 3); F9 and the status bar's "Calculate" work.

**Measured** (core benchmarks, this machine shared with other builds, load average 8–25 on 16 cores)

- 100,000-formula chain, recalculation after editing its first cell: 150–200 ms when the machine is quiet,
  200–600 ms under load. Budget 200 ms: passes in a quiet moment, fails under load.
- 50,000 VLOOKUPs over 10,000 rows, all recalculated: 490 ms (budget 1 s).
- Inserting a row above 100,000 formulas (rewrite and rebuild): 1.2 s, undo 1.5 s (budgets 2 s and 3 s).
- Editing one value under `SUM(A:A)` with 50,000 values: 14 ms.

**Open**: see `docs/verify-list.md`.

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
