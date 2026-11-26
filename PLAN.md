# Cells — Implementation Plan

Status: **proposal, awaiting approval.** No application code is written until this plan is approved. After that,
one phase at a time: each phase ends with its tests green, its Excel checks done, and an explicit go-ahead before
the next phase starts.

---

## 0. What the Inkspire repo tells us

Inkspire (the word processor, `../Inkspire`, v0.1.6) was inspected before writing this plan. The facts below decide
most of the architecture.

| Area           | Inkspire today                                                                                                                                                                                                                                           | Consequence for Cells                                                                                                 |
| -------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Shell          | Tauri 2 (Rust, edition 2024), frameless transparent window, own title bar (`title-bar.tsx`, `lib/window-chrome.ts`)                                                                                                                                      | Same: Tauri 2, same window config and title bar                                                                       |
| UI             | React 19, TypeScript 6, Vite 8, Tailwind CSS 4, shadcn + `@base-ui/react`, `lucide-react`, `zustand`                                                                                                                                                     | Same versions, same component set                                                                                     |
| Look           | Office-like ribbon (`components/ribbon/*`, `document/ribbon-layout.ts`), tokens in `src/index.css` (`--app-theme-primary`, `--app-ribbon-*`, …), light/dark                                                                                              | Same tokens with a Cells theme colour (green family instead of Word blue); same ribbon components                     |
| Files          | `lib/file-io.ts` wraps Tauri dialog/fs with a browser fallback; Rust `read_document` / `write_document` commands                                                                                                                                         | Same pattern behind a `FileSystem` port (§2)                                                                          |
| Office formats | Own OPC/zip layer (`document/docx/opc.ts`, `@zip.js/zip.js`, `saxes`), DrawingML chart writer (`chart-xml.ts`, `chartex-xml.ts`), and it already writes a small embedded `.xlsx` for every chart (`inlines.ts`)                                          | Share OPC, chart XML and chart spec code; the xlsx reader/writer becomes the real one, and Inkspire later uses it too |
| Charts         | `editor/chart-spec.ts`, `chart-sheet.ts` (series/categories model), Recharts rendering, docked data sheet dialog                                                                                                                                         | Same chart model, so a chart moves between the two apps without conversion                                            |
| Formulas       | A 4×4 "Spreadsheet block" (`editor/sheet-block.ts`) and an arithmetic-only `evaluateFormula`                                                                                                                                                             | Nothing reusable for a real engine; the Cells engine later replaces both in Inkspire                                  |
| Storage        | Local only: documents on disk, recent files / recovery drafts / autosave in app data (`document/app-data.ts`, `recovery.ts`, `autosave.ts`), API keys in the OS keyring (`secrets.rs`), user settings in webview `localStorage` (`aiword.user-settings`) | No accounts and no cloud exist to replicate (§Phase 5)                                                                |
| Tests          | Vitest + jsdom + Testing Library, `cargo test --lib`; a `.docx` corpus with round-trip tests; CI: lint, typecheck, vitest, build, cargo test                                                                                                             | Same toolchain and CI shape, plus an `.xlsx` corpus                                                                   |
| Release        | GitHub Actions: `.deb`, AppImage, Flatpak (GNOME runtime), Windows NSIS, macOS universal `.dmg`, Android tablet `.apk` built from an `android` branch                                                                                                    | Same workflow; phone layout is new                                                                                    |
| Android        | `isAndroid` in `lib/platform.ts`, `@media (pointer: coarse)` touch sizing, SAF-aware paths (`lib/paths.ts`) — only on the `android` branch                                                                                                               | Reuse; phone layout is a separate shell                                                                               |
| Reference app  | Local, git-ignored tooling renders files in the real Word inside the Office VM through the QEMU guest agent (PowerShell + COM, hidden instance, no desktop capture)                                                                                      | Same mechanism driving Excel (§4)                                                                                     |

**Placeholders in the brief (confirmed 2026-11-14).** The Word processor is Inkspire; there is no separate Inkspire
SDK, integration means matching Inkspire's local-only setup (Phase 5); the reference VM is the Office VM Inkspire's
Word tooling already uses, and Excel is installed there.

---

## 1. Decisions that shape everything

1. **Separate repo, same stack (decided 2026-11-14).** The app is named **Cells**, lives in this folder and gets its
   own GitHub repo (not inside Inkspire), developed on `main`.
   Modules taken from Inkspire are copied with a provenance note (source path + Inkspire commit). Extracting a shared
   package (`@inkspire/office-ui`, `@inkspire/ooxml`) is a Phase 5 decision, once both apps use the same code and we
   know what really is shared — not before.
2. **The engine is TypeScript, running in a Web Worker.** Same language as the UI, runs unchanged in WebKitGTK,
   WebView2, WKWebView and Android WebView, and is tested in Node with Vitest. Rust stays a thin platform layer
   (files, keyring, windows). A Rust/WASM calc core is a possible later optimisation only if profiling in Phase 2
   says so; the worker message API is designed so the engine could be swapped without touching the UI.
3. **Canvas grid, DOM chrome.** Cells are painted on a `<canvas>` (only the visible tiles); the cell editor,
   selection handles, comments and menus are DOM over it. DOM-per-cell cannot reach 100k+ rows smoothly in
   WebKitGTK.
4. **Every change is an operation.** All edits go through typed, serialisable operations
   (`setCells`, `insertRows`, `format`, `sort`, …) applied by the core. Undo/redo, macro-free replay, the worker
   protocol, autosave and any future sync all use the same log.
5. **Excel 365 semantics are the target**, including dynamic arrays and spilling, the 1900 date system (with its
   1900-02-29 quirk), 15-significant-digit display, case-insensitive text comparison, and the 1,048,576 × 16,384
   sheet bounds.
6. **Look (decided 2026-11-14).** Cells is part of the Inkspire family: **Office-like style (`office-style-ui`),
   Inkspire's tokens, own app icon** — not the FileSort house style.
7. **Original assets.** Excel's layout, behaviour and shortcuts are matched; Microsoft's icons, artwork, fonts
   (Aptos/Calibri are not bundled) and names are not used. Icons come from lucide or are drawn for the app.

---

## 2. Architecture

```
Cells/
  packages/
    core/          pure TS, no DOM, no Tauri, no React  ← the spreadsheet
      model/       workbook, sheets, sparse cell store, styles table, names, tables
      ops/         operations, apply/invert, undo stack, structural ref-rewrite
      formula/     lexer, parser → AST, evaluator, function library, error values, coercion rules
      calc/        dependency graph, dirty propagation, topological recalc, cycles, volatiles, spills
      format/      number-format parser + renderer (Excel codes, locales, dates)
      view/        viewport maths: row/col size index (Fenwick tree), hit-testing, freeze panes, merges
      clipboard/   TSV/HTML/internal-JSON encode + decode, paste-special
      data/        sort, filter, find/replace, fill series, conditional-format rules
    io/            pure TS file formats (depends on core only)
    ooxml/         pure TS zip, XML and OPC for workbooks and Word documents (shared with Inkspire, Phase 5)
      xlsx/  csv/  opc/  chart-xml/
    platform/      interfaces only: FileSystem, Dialogs, Clipboard, Window, Settings, Keyring, AppLink
  app/
    src/
      engine-host/     Web Worker wrapper around core; async API for the shells
      grid/            canvas renderer, editor overlay, pointer/touch/keyboard controllers
      shells/
        desktop/       DesktopTabletShell: title bar, ribbon, formula bar, sheet tabs, status bar, panes
        phone/         PhoneShell: compact header, collapsible formula bar, bottom sheet switcher, tool drawers
      platform-tauri/  implementations of packages/platform for Tauri (desktop + Android)
      platform-web/    browser fallback (dev server, tests)
  src-tauri/           Rust: file I/O commands, keyring, single-instance/deep links, window chrome
  tests/xlsx-fixtures/ corpus + Excel-produced expectations
  tools/               local Excel tooling (git-ignored, as in Inkspire)
```

**Boundary rules, enforced not hoped for:** ESLint `no-restricted-imports` + a dependency-cruiser check in CI so
`packages/core` and `packages/io` cannot import React, DOM types, `@tauri-apps/*` or anything in `app/`;
`packages/core` builds with `lib: ["ES2023"]` (no `DOM`), so a stray `window` fails typecheck. Core tests run in plain
Node.

**Data model essentials.**

- Sparse storage per sheet: column-major chunks of 1,024 rows; values in typed arrays where numeric, a string
  pool for text, formulas stored once as shared ASTs (R1C1-relative) like Excel's shared formulas.
- Styles are interned (one id per distinct format), cells hold the id.
- Row heights/column widths: defaults + overrides in a Fenwick tree, so "pixel offset of row 800,000" and "row at
  pixel y" are O(log n).
- Worker protocol: shells send operations, receive patches (changed cells in or near the viewport, structural
  events). The UI never holds the whole sheet.

---

## 3. Phases

Each phase lists scope, the tests that prove it, the Excel checks, and the acceptance bar. Numbers are budgets on a
mid-range Linux laptop in the release build; they are measured, not estimated, at the end of each phase.

### Phase 1 — Core grid engine and Linux desktop shell

**Scope**

1. Repo scaffold: npm workspaces, Tauri 2 app (Linux first), CI (lint, typecheck, vitest, build, cargo test,
   boundary check), commit identity and privacy checks as in your global rules.
2. Core model: workbook/sheets, sparse cell store (values: number, text, boolean, error, empty), operations,
   undo/redo stack with grouping (typing + Enter is one step; a paste is one step) and selection restore.
3. Viewport maths: virtual scrolling over 1,048,576 rows × 16,384 columns, variable row/col sizes, hidden rows/cols,
   hit-testing, scroll-into-view.
4. Canvas renderer: gridlines, headers, text with clipping/overflow into empty neighbours, alignment, HiDPI,
   tile cache, light/dark.
5. Selection: active cell, ranges, multi-range (Ctrl+click), whole rows/cols, Select All, Shift+click extend, drag
   select with auto-scroll, fill handle (copy + simple series 1,2,3 / Mon,Tue / dates).
6. Editing: type-to-replace, F2 edit in place, Enter/Tab/Shift variants and the "return to Tab-start column"
   rule, Esc cancel, Alt+Enter line break, Delete clears, Ctrl+Enter fills selection.
7. Keyboard navigation per Excel: arrows, Ctrl+arrows (data-edge jumps), Home/Ctrl+Home/Ctrl+End, PgUp/PgDn,
   Alt+PgUp/PgDn, Ctrl+Space / Shift+Space, End mode, Ctrl+Shift+arrows, F5 / Ctrl+G Go To.
8. Copy/cut/paste: TSV + HTML out, TSV/HTML/plain in, internal rich format inside the app, marching-ants cut, paste
   to a larger selection tiles; system clipboard through the `Clipboard` port.
9. Insert/delete rows and columns, resize and auto-fit (double-click border), hide/unhide.
10. Minimal shell: title bar, a placeholder formula bar (shows/edits raw text), sheet canvas, status bar
    (Sum/Average/Count of selection for numbers). The ribbon comes in Phase 3.
11. Local Excel tooling skeleton (`tools/excel-run.py`, §4) so Phase 2 can start verifying immediately.

**Tests:** core unit tests for every operation and its inverse (property test: random op sequences, undo all,
compare to start); viewport maths property tests; keyboard-navigation table tests (start cell + data layout + key →
expected cell, the table filled from Excel); clipboard round-trips; renderer smoke tests with a canvas mock; one
Playwright/WebDriver run of the Tauri app on Xvfb for scroll + edit + undo.

**Excel checks:** navigation and selection tables recorded from Excel (Ctrl+arrow over gaps, End mode, Tab/Enter
return column, fill-handle series), clipboard TSV/HTML that Excel produces and accepts.

**Acceptance:** 100,000 populated rows × 20 columns: open < 1 s, scroll at 60 fps (frame time p95 < 16 ms), jump to
the last row < 50 ms; 1,000-step undo history works without growth problems; runs as a Linux `.deb`/AppImage.

### Phase 2 — Formula and calculation engine

**Scope**

1. Lexer/parser for Excel grammar: A1 and `$A$1` mixed refs, ranges, whole rows/cols (`A:A`, `3:3`), cross-sheet
   (`Sheet2!A1`, `'My Sheet'!A1:B2`), 3-D refs (`Sheet1:Sheet3!A1`), defined names, array constants `{1,2;3,4}`,
   operator precedence including `%`, `^`, `&`, unary minus vs `^` (Excel's `-2^2 = 4`), range `:`, union `,`,
   intersection ` `, implicit intersection `@`, structured refs as a stub for Phase 4 tables. R1C1 parse/print for
   shared formulas and the R1C1 option.
2. Values and errors: number, text, boolean, error (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`,
   `#N/A`, `#SPILL!`, `#CALC!`), empty-vs-zero, Excel's coercion rules (text "1" in `+` vs in `SUM`), floating-point
   display at 15 digits and Excel's near-equality comparisons.
3. Dependency graph: cell→cell edges plus range dependencies indexed by column bands (so `SUM(A:A)` does not create a
   million edges), dirty propagation, topological recalc, circular-reference detection + status-bar warning,
   optional iterative calc, volatile functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`),
   manual/automatic calc mode, F9 / Shift+F9.
4. Dynamic arrays: spill ranges, `#SPILL!` on collision, `A1#` spill references, implicit intersection for legacy
   formulas.
5. Structural edits rewrite formulas: insert/delete rows/cols/sheets shift refs and produce `#REF!`; moving cells
   (cut/paste, drag) rewrites references to them; copy adjusts relative refs.
6. Function library, first wave (≈150), grouped:
   - Math: `SUM SUMIF SUMIFS SUMPRODUCT PRODUCT ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD ABS SIGN SQRT POWER EXP LN LOG
LOG10 PI RAND RANDBETWEEN CEILING(.MATH) FLOOR(.MATH) MROUND EVEN ODD GCD LCM FACT COMBIN SUBTOTAL AGGREGATE
SEQUENCE RANDARRAY` + trig.
   - Statistics: `AVERAGE AVERAGEIF(S) COUNT COUNTA COUNTBLANK COUNTIF(S) MIN MAX MINIFS MAXIFS MEDIAN MODE.SNGL
LARGE SMALL RANK.EQ STDEV.S STDEV.P VAR.S VAR.P PERCENTILE.INC QUARTILE.INC CORREL`.
   - Logical: `IF IFS IFERROR IFNA AND OR XOR NOT TRUE FALSE SWITCH LET LAMBDA` (LAMBDA may slip to a later wave).
   - Lookup/reference: `VLOOKUP HLOOKUP XLOOKUP XMATCH LOOKUP INDEX MATCH CHOOSE OFFSET INDIRECT ROW(S) COLUMN(S)
ADDRESS TRANSPOSE FILTER SORT SORTBY UNIQUE TAKE DROP HSTACK VSTACK`.
   - Text: `CONCAT CONCATENATE TEXTJOIN LEFT RIGHT MID LEN FIND SEARCH SUBSTITUTE REPLACE UPPER LOWER PROPER TRIM CLEAN
TEXT VALUE NUMBERVALUE REPT EXACT CHAR CODE UNICHAR UNICODE TEXTBEFORE TEXTAFTER TEXTSPLIT`.
   - Date/time: `DATE TIME DATEVALUE TIMEVALUE TODAY NOW YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY WEEKNUM ISOWEEKNUM
EDATE EOMONTH DATEDIF NETWORKDAYS(.INTL) WORKDAY(.INTL) YEARFRAC DAYS`.
   - Information: `ISBLANK ISNUMBER ISTEXT ISERROR ISERR ISNA ISLOGICAL ISFORMULA ISREF N NA TYPE ERROR.TYPE`.
   - Financial: `PMT IPMT PPMT FV PV NPER RATE NPV IRR XNPV XIRR`.
     Criteria strings (`">=10"`, wildcards `* ? ~`) are one shared module used by every *IF/*IFS function.
7. Formula bar wired to the engine: shows the formula, edits it, colour-coded reference highlighting on the grid,
   point mode (arrow keys/click insert refs while typing), F4 cycles `$` styles, function autocomplete with argument
   tooltips, Ctrl+` show formulas, Ctrl+Shift+Enter accepted for legacy arrays.

**Tests:** parser golden tests (formula → AST → printed formula is stable); a **conformance suite**: thousands of
`(inputs, formula) → value` cases generated and evaluated by Excel in the VM, stored as JSON in
`tests/excel-golden/`, replayed by Vitest; structural-edit tests (insert/delete/move → expected rewritten formulas,
recorded from Excel); recalc ordering and cycle tests; performance benchmarks in CI (fail on >20 % regression).

**Excel checks:** every function wave is checked against the conformance suite; edge cases recorded explicitly —
empty-cell coercion, text-number coercion, error precedence (`#N/A` vs `#VALUE!` in the same formula), rounding of
`ROUND(2.675,2)`, `MOD` with negatives, date serial 60, `VLOOKUP` approximate-match on unsorted data, `XLOOKUP`
wildcard/binary modes, `MATCH` type −1.

**Acceptance:** 100 % pass on the conformance suite for implemented functions (known differences documented one by
one); recalc of a 100,000-formula dependency chain < 200 ms; a 50,000-row sheet with `VLOOKUP` over 10,000 rows
recalculates < 1 s; unknown functions show `#NAME?` but keep their text so files round-trip.

### Phase 3 — Full desktop/tablet UI and formatting

**Scope**

1. Ribbon, using Inkspire's ribbon components and layout engine (groups collapse as the window narrows):
   **File (backstage)**, **Home**, **Insert**, **Page Layout**, **Formulas**, **Data**, **Review**, **View**, **Help**;
   contextual tabs later (Table Design in Phase 4, Chart Design/Format in Phase 4). Quick Access Toolbar, Tell Me /
   search box, KeyTips (Alt then letters) matching Excel's letters, collapse ribbon (Ctrl+F1).
2. Formula bar (full): Name Box with go-to and named ranges, fx button → Insert Function dialog, expand/collapse
   multi-line, resizable.
3. Sheet tabs: add, rename (double-click), move/copy (drag, Ctrl+drag), colour, hide/unhide, scroll arrows,
   Ctrl+PgUp/PgDn, right-click menu.
4. Cell formatting: font, size, bold/italic/underline/strikethrough, colours, fill, alignment (h/v, indent, wrap,
   shrink to fit, rotation), merge/unmerge/merge across/center across selection, borders (all Excel styles and the
   Borders dropdown + Format Cells Border tab), cell styles gallery, Format Painter, clear formats.
5. Number formats: a full Excel format-code parser/renderer (sections, conditions, colours, `0 # ? , . % E+`, text
   `@`, dates/times/elapsed `[h]`, fractions, locale currency), the Number Format dropdown, Increase/Decrease Decimal,
   Format Cells dialog (Number, Alignment, Font, Border, Fill, Protection tabs), auto-detection when typing
   (`10%`, `1/2`, dates in the user's locale, `$1,000`).
6. Conditional formatting: highlight rules, top/bottom, data bars, colour scales, icon sets, formula rules, Manage
   Rules dialog with priority and Stop If True.
7. View: freeze panes (top row, first column, at selection), split panes, zoom (status bar slider, Ctrl+wheel,
   pinch), gridlines/headings toggles, Normal/Page Break Preview/Page Layout (Page Layout view may move to Phase 4),
   multiple windows on one workbook, dark mode.
8. Review basics: notes and threaded comments (local author name), protect sheet/workbook (no encryption yet).
9. Context menus for cells, headers, tabs; Excel's full shortcut set for everything in this phase (Ctrl+1, Ctrl+B/I/U,
   Ctrl+Shift+$ % # @ ! ~ ^ &, Ctrl+Shift+_ , Alt+H shortcuts via KeyTips, Ctrl+; / Ctrl+Shift+; etc.).
10. Tablet behaviour in the same shell: `pointer: coarse` sizing (from Inkspire's Android branch), touch selection
    handles, pinch zoom, long-press context menu, on-screen keyboard aware editor, stylus as mouse.
11. App icon: own designed SVG master and all platform sizes, matching the in-app mark.

**Tests:** number-format conformance (format code × value → text, generated by Excel's `TEXT` in the VM); component
tests for every dialog; ribbon command tests (command id → operation emitted); conditional-format evaluation tests;
screenshot tests of the shell in light and dark on Xvfb with mock data only.

**Excel checks:** side-by-side screenshots of each ribbon tab and dialog against Excel at the same window size;
number-format output for a few hundred codes; conditional-format results on fixture sheets; shortcut list cross-checked
with Excel's.

**Acceptance:** every control on the listed tabs either works or is visibly disabled with a tooltip saying it is not
built yet (no dead buttons); formatting a 100k-row column < 300 ms; frozen panes scroll at the same frame budget as
Phase 1.

### Phase 4 — File I/O and data tools

**Scope**

1. `.xlsx` read/write in `packages/io` (streaming SAX via `saxes`, zip via `@zip.js/zip.js`, Inkspire's OPC layer):
   values, shared strings, shared/array/dynamic formulas, styles (fonts, fills, borders, number formats, cell styles,
   theme colours), column widths/row heights, merges, freeze/split panes, sheet visibility and colour, defined names,
   conditional formatting, data validation, hyperlinks, notes/comments, tables (`ListObject`), autofilter, print
   settings, charts. **Unknown parts are preserved byte-for-byte** and written back (pivot caches, VBA project in
   `.xlsm` kept as opaque data, slicers, drawings we do not draw), following the rules in the `office-docx-fidelity`
   skill as applied to SpreadsheetML.
2. Large files: reading in the worker with progress; a 50 MB / 1 M-cell `.xlsx` opens without freezing the UI.
3. `.csv` / `.tsv` / `.txt`: import wizard (delimiter, encoding detection with UTF-8 BOM, quote rules, column types,
   locale-aware numbers/dates), export with chosen delimiter/encoding. Optional later: `.ods` through the same
   preserving layer.
4. Platform-abstracted dialogs: `Dialogs.openFile/saveFile` implemented by Tauri dialogs on desktop, the Android
   system picker (SAF) on Android, the File System Access API in the browser dev build; recent files, autosave and
   crash recovery as Inkspire does them.
5. Data tools: sort (multi-level dialog, by value/colour, case option, header detection), AutoFilter (value list
   with search, number/text/date filters, top 10, by colour), Format as Table with table styles and structured
   references, remove duplicates, text to columns, data validation (lists with dropdown), Flash Fill (stretch), find
   and replace (within sheet/workbook, formulas/values, match case/entire cell), Go To Special.
6. Charts: Insert Chart (column, bar, line, pie/doughnut, area, scatter, combo) using Inkspire's chart spec and
   renderer, chart elements/styles, Chart Design and Format contextual tabs, read/write as DrawingML chart parts
   (Inkspire's `chart-xml.ts`), sparklines (stretch).
7. Printing/PDF: page setup, print area, titles, headers/footers, Page Layout view, export to PDF.

**Status (2026-11-15):** items 1–7 are built and committed. Left out for later: Page Layout view and Page Break
Preview, sparklines and the chart Format tab (`.ods` came in Phase 8). PDF pages are images (with searchable text since Phase 8). What's left to check
by hand is in `docs/verify-list.md`, and the remaining steps are in `docs/next-todo.md`. Phase 4 is done; the summary
is in `docs/changes.md`.

**Tests:** an `.xlsx` corpus (made-up data only) with round-trip tests: open → save → re-open must equal, and the
saved file must open in Excel without a repair prompt; unknown-part preservation tests; a schema validator pass on
every saved file (Open XML SDK validator, as Inkspire's `tools/docx-validate`); CSV edge cases (embedded quotes,
newlines in fields, `;` locales, leading zeros); sort/filter tests against Excel's results (stable sort, blanks last,
mixed types order).

**Excel checks:** every corpus file is opened by Excel in the VM after our save (no repair, values and formats match
a cell-by-cell dump taken from Excel); files saved by Excel open in Cells with matching values and formats;
screenshot/PDF comparison for charts and printed pages.

**Acceptance:** 0 repair prompts across the corpus; 100 % cell value/formula equality after round-trip; formatting
differences listed and justified per file.

### Phase 5 — Inkspire ecosystem integration

Inkspire has **no account login, no cloud storage and no sync service**: documents are local files, app state lives in
the app-data folder, API keys in the OS keyring. So "the same authentication and storage" means matching that
local-first model; there is nothing to log in to (confirmed 2026-11-14). Building a cloud service is out of scope.

**Scope**

1. Shared identity and settings: the user's name (author of comments), regional formats and theme are kept today in
   Inkspire's webview `localStorage`, which another app cannot read. Move them to a shared settings file in a
   common config folder (`$XDG_CONFIG_HOME/inkspire/` and platform equivalents) that both apps read/write, with
   Inkspire migrating its existing value. Same for the recent-files list (shown in both apps' File views) if wanted.
2. Shared keyring entries: if Cells gains AI tools, it reads the same keyring service Inkspire's `secrets.rs` uses,
   so keys are entered once.
3. App linking: both apps register a URL scheme and become single-instance (Tauri `deep-link` +
   `single-instance` plugins): `cells://open?file=…&range=Sheet1!A1:D10`, `inkspire://open?file=…`.
   "Edit in Cells" from Inkspire and "Open in Inkspire" from Cells use it.
4. Embedding (static): copy a range or chart in Cells → paste in Inkspire as a formatted table or an editable chart.
   The clipboard carries HTML + TSV (works with any app) and an Inkspire MIME type with full fidelity (styles, the
   chart spec, the source workbook and range).
5. Embedding (live/linked), in Word's own file format so it also works in Word:
   - linked table: a `LINK Excel.Sheet.12 "file.xlsx" "Sheet1!R1C1:R10C4" \a \f 4 \h` field (Inkspire already has a
     field engine), refreshed on open, on demand ("Update Link"), and live while both apps run (Cells notifies
     Inkspire through the app link channel when a linked range changes);
   - linked chart: chart part with `c:externalData` pointing to the external workbook;
   - embedded chart: the chart's data lives in the embedded `.xlsx` Inkspire already writes; "Edit Data" opens that
     workbook in Cells instead of Inkspire's small data-sheet dialog, and the result is written back.
6. Shared code extraction (decided here, not earlier): move OPC/zip, chart spec/XML, ribbon components, tokens and the
   formula engine into shared packages; Inkspire's 4×4 Spreadsheet block and arithmetic formula fields then use the
   real engine.
   **Decided 2026-11-16** (`docs/shared-packages.md`): `@cells/ooxml` (new: zip, XML, OPC for workbooks and Word
   documents), `@cells/chart` and `@cells/core` (with `calculateGrid`/`calculateFormula`) are shared from this repo;
   Inkspire takes them as a submodule and npm workspaces. Ribbon components and tokens are not shared: the two ribbons
   have no code in common, and only 19 token names match.

**Decisions (2026-11-15):** Cells gets AI features in this phase, using Inkspire's AI providers and the same keyring
entries (opt-in). The regional formats are applied in Cells, not only stored. Each app keeps its own recent-files list
(item 1's "if wanted" is no). The shared file's format is in `docs/shared-settings.md`.

**Required changes in Inkspire** are deferred and tracked in `docs/inkspire-todo.md`; they are done later, in
the Inkspire project, not from this repo. Cells' side of each item is built so it works on its own until Inkspire
catches up.

**Tests:** settings migration tests; deep-link parsing tests; a cross-app end-to-end on Xvfb (open a doc with a
linked range, change the cell in Cells, see it update in Inkspire); Word in the VM opens Inkspire's files with the
linked table/chart and can update the link; Excel opens the embedded workbook.

**Acceptance:** a chart or range made in Cells can be pasted, linked, updated and edited from Inkspire, and the
resulting `.docx` behaves the same way in Word.

### Phase 6 — Cross-platform packaging and phone UI

**Decisions (2026-11-16):** the platforms are as below, with no iOS/iPadOS. Speed is measured in the Android
emulator for now; the real-phone measurement in the acceptance waits until a device is named (`verify-list.md`).
Android signing: test builds use the debug key; the release job reads Inkspire's secret names and key alias
(`ANDROID_KEY_BASE64`, `ANDROID_KEY_PASSWORD`, alias `inkspire`), so Inkspire's key can be reused by copying its two
secrets into this repo later. Until then the release `.apk` is left unsigned.

**Scope**

1. Desktop packaging with Inkspire's release workflow: Linux `.deb`, AppImage, Flatpak; Windows NSIS `.exe`; macOS
   universal `.dmg`; file associations for `.xlsx`/`.csv`; macOS menu bar and `⌘` shortcuts mapped from the Excel
   for Mac set.
2. Android build (Tauri mobile), from `main` with the generated Android project committed (no separate
   `android` branch): tablet `.apk`/`.aab` signed in CI; SAF file access (Inkspire's `paths.ts`), share
   intent "Open with", hardware keyboard shortcuts, stylus.
3. Shell choice by screen, not by OS: `PhoneShell` when the smallest window dimension is under ~600 dp _and_ the
   pointer is coarse; otherwise `DesktopTabletShell`. An Android tablet gets the full ribbon; a phone in landscape
   still gets the phone shell; split-screen on a tablet can switch live (state lives in the engine, so a switch loses
   nothing).
4. PhoneShell: compact top bar (file name, undo/redo, more), grid with finger-sized headers and drag handles for
   selection, a formula bar that collapses to a single line and expands into a full editor with a formula keyboard
   row (`= ( ) , : $ + - * /` and a function picker), bottom bar with sheet switcher (sheet list in a bottom sheet),
   tool drawers replacing the ribbon (Format, Insert, Data, Formulas, View) as bottom sheets with gesture
   dismissal, long-press context menus, pinch zoom, double-tap to edit, inertia scrolling that stays at the frame
   budget.
5. Phone-specific limits stated up front: dialogs become full-screen sheets; features that need a mouse (e.g.
   drawing chart elements precisely) open in a simplified form.

**Tests:** shell-selection tests across a size/pointer matrix; component tests of PhoneShell drawers and editor;
emulator smoke tests (Android emulator in CI on x86_64) for open/edit/save; screenshot tests at phone and tablet
sizes with mock data; the same engine test suite runs unchanged (proof that nothing engine-side forked).

**Acceptance:** release workflow produces all installers from one tag; the phone build edits, saves and round-trips
a corpus file; scroll and edit meet the frame budget on a mid-range Android phone (measured on a real device —
please name one).

---

### Phase 7 — The commands marked "not built" (0.7.0, 0.7.1) ✅

The checklist, with what each part covers, is `docs/next-todo.md`; PivotTables have their own plan in
`docs/pivot-plan.md`.

- ✅ Parts 1–9: menus wired up, formula auditing, outline and Subtotal, data tools, links and symbols, views and
  windows (as panels inside Cells), Review tools, objects (shapes, pictures, themes, equations, 3D models, SmartArt,
  signature lines, embedded files), sparklines, slicers and Quick Analysis.
- ✅ Part 10: PivotTables, PivotCharts, slicers and Timelines for PivotTables, Recommended PivotTables, and
  "Describe the PivotTable you want" (AI, opt-in), with sorting by a value.
- ✅ Feature hints ("New: …", the house rule): `app/src/hints/`, one at a time beside the feature's control, on by
  default, File › Options › General turns them off or shows them again; kept in `apps.cells` of the shared settings.
  Each new special feature adds its hint to `feature-hints.ts`.
- ✅ Help tab (`app/src/help/`): Help (F1) is a pane with the special features (where each is, "Show me" pointing
  at its button with the hint callout) and the keyboard shortcuts, searchable and offline; Show Training walks through
  the features; Contact Support and Feedback open a pre-filled GitHub issue the person reviews and submits (Cells
  sends nothing, no workbook content). File › Help and File › Feedback too (phones have no Help tab).
- ⬜ Next: Inkspire's side of the link (`docs/inkspire-todo.md`, I1–I9), done in the Inkspire project.
- ⏸ Later: the lists under "Later" in `docs/next-todo.md` (among them: Field Headers, new PivotTable
  styles).

**Known limits:** Help's
shortcut list is written by hand (`help-content.ts`): a new shortcut has to be added there too. Show me can't point
on phones (the pane fills the screen), and Show Training says so when a feature's button isn't on screen.

### Phase 8 — Closing the gap with Excel (0.8.0, 0.8.1) ✅

In order: the functions Excel has and Cells didn't (about 250), then the greyed-out ribbon buttons, then the thinner
features (chart Format tab, Insert Chart dialog, Move Chart and chart sheets, Page Layout view and Page Break
Preview, picture crop and corrections, PivotTable extras, searchable PDF text, notes that move with inserted cells,
Name Box lists, live Flash Fill, .xlsb and .ods).

- ✅ Functions, wave 1: statistics (87): the normal, lognormal, t, chi-squared, F, beta, gamma, exponential, Weibull,
  binomial, Poisson, hypergeometric and negative binomial distributions and their inverses (new and older names),
  GAMMA, GAMMALN, CONFIDENCE, FISHER, PERMUTATIONA, PHI, GAUSS, STANDARDIZE; Z.TEST, T.TEST, F.TEST, CHISQ.TEST,
  KURT, SKEW, TRIMMEAN, PERCENTRANK, PROB, MODE.MULT, FREQUENCY; LINEST, LOGEST, TREND, GROWTH
  (`functions/distributions.ts`, `stat-tests.ts`, `regression.ts`, special functions in `special.ts`). 215 Excel
  conformance cases; the xlsx `_xlfn.` list now holds every name Excel prefixes.
- ✅ Wave 2: engineering (54): number bases, BIT*, ERF/ERFC, DELTA, GESTEP, the Bessel functions (Excel's own
  approximations), CONVERT with every unit and prefix Excel knows, COMPLEX and the IM* complex numbers written as
  Excel writes them (`functions/engineering.ts`, `complex.ts`); a new Engineering category (More Functions). 940
  Excel cases, including one per unit.
- ✅ Wave 3: financial (40): coupon dates (COUP*), PRICE, YIELD, DURATION, MDURATION, ACCRINT(M), the discount and
  at-maturity securities, TBILL*, odd first and last periods (ODDF*, ODDL*), DB, DDB, VDB, AMORDEGRC, AMORLINC,
  CUMIPMT, CUMPRINC, DOLLARDE/FR, FVSCHEDULE, ISPMT, MIRR, PDURATION, RRI (`functions/bonds.ts`,
  `depreciation.ts`), with Excel's own day counts and fractional-period rules worked out from 1384 Excel cases.
- ✅ Wave 4: database (12): DSUM, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DVAR,
  DVARP with Advanced Filter's criteria ranges, computed criteria included (`functions/database.ts`); a Database
  category in Insert Function. 262 Excel cases.
- ✅ Wave 5, part 1: LAMBDA (called in place, through LET names (recursion included) and as named functions in Name
  Manager), MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, ISOMITTED; built-in functions passed by name (=BYROW(A1:C3,
  SUM)); saved with Excel's `_xlpm.` prefixes (`functions/lambda.ts`, a `call` node in the parser). 88 Excel cases
  and an Excel-made file.
- ✅ Wave 5, part 2: GROUPBY, PIVOTBY, PERCENTOF, TRIMRANGE (`functions/groupby.ts`), built from Microsoft's
  documentation with our own tests (the Excel build used for checking doesn't have them); `_xleta.` for functions
  passed by name in xlsx.
- ✅ Wave 6: text, information, matrix and web (39 functions): REGEXTEST/REGEXEXTRACT/REGEXREPLACE, ARRAYTOTEXT,
  VALUETOTEXT, ENCODEURL, FILTERXML, BAHTTEXT, PHONETIC, ASC, DBCS, the *B text functions, USDOLLAR, CELL, INFO,
  MUNIT, MDETERM, MINVERSE, SERIESSUM, MULTINOMIAL, ACOT, ACOTH, COTH, CSCH, SECH (`matrix.ts`, `text-more.ts`,
  `cell-info.ts`, `filterxml.ts`); 203 Excel cases; a Web category in Insert Function and More Functions. DOLLAR with
  negative decimals now rounds to tens/hundreds as Excel does.
- ⏸ Not planned: WEBSERVICE and IMAGE (need the network during calculation), JIS (Japanese Excel only).
- ✅ The greyed-out ribbon buttons (released in 0.8.0; Move Chart comes with chart sheets):
  - ✅ PivotTable Analyze › Field Headers (showHeaders in files), Drill Down / Drill Up (down: only the active item,
    opened to the field inside; up: the field above shows every item again).
  - ✅ Ink: a Draw tab (Select, Eraser, two pens and a highlighter with colour and thickness, Draw with Touch) and
    Review › Hide Ink (Hide Ink, Delete All Ink on Sheet, Delete All Ink). Ink is a shape with strokes (moves,
    sizes and undoes like one); strokes drawn within 2 s join one ink object, as in Excel. xlsx: Excel's own form
    (xdr:contentPart + InkML customXml part), checked in Excel 365 (opens, keeps and draws it); Excel's difference-
    encoded traces are read. Feature hint "New: Draw on the sheet".
  - ✅ Review › Unshare Workbook: on for legacy shared workbooks (the file's revision history is kept while shared,
    dropped on save once unshared; checked in Excel 365).
  - ✅ View › Macros: lists an .xlsm's or .xls's VBA macros as Excel names them and shows their code (read-only;
    MS-OVBA decompression). Cells doesn't run VBA, so Run stays off; the project stays in the file.
  - ✅ Page Layout › Effects: Office's 15 effect schemes (taken from Office's .eftx files), kept in the theme and
    written as its format scheme. Shape Format › Shape Styles: Excel's 6 × 7 theme styles; the lower rows give a
    shape a theme effect (Excel's effectRef), drawn from the workbook's effects (shadows and glows; bevels,
    reflections and 3-D show in Excel only; the style's gradient fills are drawn solid). Checked in Excel 365.
  - ✅ Insert › Screenshot: taken with the system's own picker (Linux: the XDG desktop portal's interactive
    screenshot; macOS: screencapture -i; Windows and browsers: the webview's screen picker), inserted as a picture.
    Off on Android. Not yet tried by hand on Windows or macOS (can't be built here); Excel's gallery of open windows
    is the system picker instead.
  - ✅ Data › Get & Transform Data (Power Query): From Table/Range and From Web (CSV, JSON, or a web page's tables
    through a Navigator) open a Query Editor with a preview and Applied Steps: choose/remove columns, keep first rows,
    remove duplicates, first row as headers, sort, filter rows and data types. Close & Load makes a table on a new
    sheet (or a connection only); Refresh, Refresh All and the Queries & Connections pane (Edit, Refresh, Delete) run
    them again. Recent Sources (kept on the device) and Existing Connections. Queries are kept in the file as
    Excel's DataMashup with Excel's own M, so Excel reads, edits and refreshes them (checked in Excel 365: a query
    made in Cells refreshes there; an Excel query changed in Cells too); queries Cells can't run (other M steps or
    sources) stay untouched and Excel refreshes them. Also fixed: Power Query tables used to lose their query table
    part, which Excel called a damaged file. The desktop and Android apps fetch web data themselves; the browser
    build only reads sites that allow it (CORS).
  - ✅ Fixed: a workbook with Excel's Data Model (or any connection Excel stamped with xr16:uid) wouldn't open in
    Excel after Cells saved it (connections.xml lost its namespace declarations). The model (xl/model/item.data)
    and its relationships are kept as Excel wrote them. Checked in Excel 365.
  - ✅ Data › Relationships and Manage Data Model show Excel's Data Model (read-only, the user's choice): Manage
    Relationships lists them (New/Edit/Delete off); the Data Model window has a Diagram View (tables, their
    columns, a line per relationship, keys marked) and a Data View (each table's rows from its workbook table).
    Excel trusts only its binary model (xl/model/item.data, an Analysis Services backup): without it, it opens an
    empty model even when workbook.xml lists the tables and relationships, and with it, it ignores the XML; so
    Cells can't change the model, and keeps it as Excel wrote it. Also: disabled dialog buttons now look disabled.
  - ✅ Move Chart (with chart sheets, below).
- ✅ Then the thinner features above (released in 0.8.1, with release notes):
  - ✅ Chart sheets and Move Chart: chart sheets in a file are read and saved (they used to be dropped), and show
    their chart filling the window with the chart tabs open. Chart Design › Move Chart puts a chart on a new chart
    sheet (named Chart1, …) or onto another worksheet; F11 makes a chart of the selection on a new chart sheet.
    Undo covers both. Checked in Excel 365: Excel's own chart sheet survives a round trip, and chart sheets made in
    Cells open there as chart sheets with their series. Hint and Help topic "A chart on a sheet of its own".
    Limits: the chart sheet's cell-based ribbon buttons (Home, Formulas…) stay on and do nothing useful; printing
    and PDF of a chart sheet aren't done yet; a chart sheet holds one chart (Excel allows more on top).
  - ✅ Chart Format tab: click a chart, then a part (title, axis, series, legend…) to pick it, or choose it in Chart
    Elements; it's boxed. Shape Fill, Shape Outline (colour and weight), Text Fill and Reset to Match Style change just
    that part; Format Selection or a double-click opens the Format pane (fill, border or line, font, and the part's
    options: title text, axis bounds and number format, legend position, label contents, gap width and overlap,
    slice angle and explosion, smoothing and markers). Arrange and Size as for shapes. The chart XML now writes and
    reads back every part's text size and bold, line widths and the plot area's border, keeping only what differs
    from the style (Excel's own charts gain no overrides). Checked in Excel 365: each value shows there as set, and
    Excel's resave comes back the same. Hint and Help topic "Format any part of a chart".
    Limits: no Shape Styles or WordArt galleries, Shape Effects, gradients or Text Outline (Cells' charts don't draw
    them); value-axis lines, legend and title fills aren't drawn, so they aren't offered.
  - ✅ Insert Chart dialog (Recommended Charts, the Charts launcher, each gallery's "More … Charts...", and Chart
    Design › Change Chart Type): Recommended Charts lists the types that suit the data (scatter for x-y pairs, line
    for long or dated runs, pie for one small series, a combo with a second axis for series of very different sizes),
    each with a preview of the selected data and Excel's line on what it's for; All Charts has every type by family.
    Limits: no Templates or custom Combo (series by series) page; Maps aren't offered (Cells has no map charts).
  - ✅ Page Layout view and Page Break Preview, the gaps (both views were built in 0.7.0): headers and footers are
    typed on the page, as in Excel (a click in a page's top or bottom margin opens its left, center and right boxes;
    Insert › Header & Footer opens the active cell's page), with a Header & Footer tab (Excel's ready-made headers and
    footers, Page Number, Number of Pages, Date, Time, File Path, File Name, Sheet Name, Go to Header/Footer); the
    ruler setting is kept per sheet and saved in Excel files (showRuler); Page Break Preview's right-click menu has
    Insert/Remove Page Break, Reset All Page Breaks, Set/Reset Print Area and Page Setup; a break dragged further
    than a page holds lowers the print scale so the page holds it, as Excel does. Hint and Help topic "Type headers
    and footers on the page". Limits: print titles aren't repeated on each page in Page Layout view (they are when
    printing); no Picture element or Options group (different first, odd and even pages; Excel's are kept as read).
  - ✅ Picture crop (Picture Format › Size › Crop): Crop shows black handles on the picture's sides and corners; they
    cut it (Shift on a corner keeps the proportions, outward adds empty space), a drag inside slides the picture
    under the frame, and the round handles on the faint whole picture size it. Esc, Enter, a double-click or a click
    away ends cropping; each drag is one undo step, also on flipped and turned pictures. The arrow under Crop has
    Crop to Shape (every shape in the Shapes gallery but lines; the border follows the shape), Aspect Ratio (Excel's
    1:1, portrait and landscape ratios, from the middle) and Fill / Fit (in the picture's own proportions). Reset
    Picture also undoes the shape. Saved as Excel does (srcRect, negative for Fit; prstGeom); checked in Excel 365:
    no repair, the crops, Oval and Heart read back there as set and draw the same, and Excel's resave reads back
    the same in Cells. Hint and Help topic "Crop a picture, or cut it to a shape". Also fixed: a ribbon menu command
    now leaves the keyboard on the sheet (it went back to the menu's button, so Esc and arrows did nothing).
    Limits: no Format Picture › Crop fields (picture and crop position in inches); 3D models, signature lines and
    embedded objects' icons can't be cropped (Excel doesn't crop the first two either).
  - ✅ Picture corrections (Picture Format › Adjust): Corrections (Sharpen/Soften and a 5 × 5 Brightness/Contrast
    gallery), Color (Saturation, Color Tone, Recolor: Grayscale, Sepia, Washout, Black and White 25/50/75% and the
    Dark and Light variations of the theme colours; Set Transparent Color, a click on the picture), Transparency
    (0–95%), each gallery showing the selected picture at each setting, plus Picture Corrections / Color /
    Transparency Options (a slider and a number per setting). Reset Picture is now a menu: Reset Picture (look) and
    Reset Picture & Size. Corrections are baked into the picture as Excel does, from the original kept beside it (a
    PNG; Excel's own may be an HD Photo): brightness, contrast and colour tone with Excel's own curves (measured in
    Excel 365), saturation exactly, sharpen and soften close. Colour and transparency are drawn as Office 2007 draws
    them (formulas fitted to Excel's PDF output), so they print and export too. Saved as Excel does (a:grayscl,
    a:biLevel, a:duotone, a:lum, a:alphaModFix, a:clrChange; a14:imgProps with the original as an hdphoto layer);
    checked in Excel 365: no repair, Excel reads every setting back (PictureEffects, ColorType, brightness and
    contrast, the transparent colour), and Excel's resave reads back the same in Cells. Hint and Help topic
    "Corrections, Color and Transparency for pictures". Limits: a picture whose original Excel kept as an HD Photo
    keeps its corrections as they are (colour and transparency still change); Set Transparent Color takes the exact
    colour (as Excel); icons and SVG pictures can't be adjusted.
  - ✅ Artistic Effects (Picture Format › Adjust): Excel's 22 effects in a gallery of the selected picture, and
    Artistic Effects Options (the effect, its transparency and its own setting, with Excel's ranges and defaults).
    Drawn as look-alikes of Excel's (Office's filters aren't published; compared side by side with Excel 365's on a
    mock photo), at the size Excel works at (the picture's size on the sheet at 220 dpi), and saved as Excel does
    (a14:artistic…, defaults left out, in Office's order among the corrections). Checked in Excel 365: no repair,
    Excel reads each effect, its settings and their order back, and redraws it in its own way once it's changed
    there; Excel's own effects read and save back the same. Hint and Help topic "Artistic Effects for pictures".
    Limits: the look is close, not Excel's pixels; one effect at a time (as Excel's ribbon; a file with two keeps
    the second as it was, and its corrections then can't change in Cells); a resized picture keeps the effect at the
    size it was made for until it's changed.
  - ✅ PivotTable extras, in parts:
    - ✅ Show Report Filter Pages (PivotTable Analyze › Options ▾): a dialog lists the report filters; OK makes a
      sheet for each item (only the ticked ones when several can be ticked), named after the item as Excel names it
      (an item that can't name a sheet gets "SheetN", a name taken gets " (2)"), placed before the PivotTable's sheet
      in the filter's order, each with a copy of the PivotTable at A1 showing that item; one undo step. Excel 365's
      own run on the same data gave the same sheets, names, order and PivotTable names; Cells' result opens in Excel
      without repair and refreshes to the same cells. Hint and Help topic.
    - ✅ Sorting by a value: Sort on a PivotTable value (cell menu, Data › Sort A to Z / Z to A) orders the field of
      its row by that value (of its column in the Grand Total row), as Excel's; More Sort Options (cell menu,
      a field's dropdown, Data › Sort) is Excel's Sort (Field) dialog: Manual (keeps the order shown), Ascending or
      Descending by the field's labels or a value field, with its summary line. Checked in Excel 365: no repair,
      Excel reads each sort (field, direction, value field) and refreshes to the same cells. Hint and Help topic.
      Limits: by the grand totals only (Excel can also sort by the values of one column, and has a More Options
      dialog for that and for first-key sort order); a value under one column item says so instead of sorting.
    - ✅ Calculated items (Fields, Items, & Sets › Calculated Item): an item worked out by a formula over its field's
      other items (=Pens+Ink*2; quoted names, Field[Item], places like Product[1], worksheet functions), shown under
      every outer item, shared by the PivotTables on the same data. Worked out as Excel 365's, measured: the formula
      in the innermost cells on the value field's figures (sum, count, max…), every total putting its cells together
      with the value field's function; where two meet, the one last in Solve Order. Solve Order (Move Up / Down,
      Delete) and List Formulas (Excel's sheet: calculated fields and items with their solve order, and its note).
      Excel's rules kept: no Average, StdDev or Var and no field summarized twice; only items of the same field;
      not on grouped fields. Read and written as Excel does (sharedItems f="1", calculatedItems in solve order); a
      PivotTable with calculated items used to be kept as it was and is now read. Checked in Excel 365: no repair,
      Excel lists the formulas in the same solve order, and refreshes three PivotTables (crossing items, Count and
      Max, nested fields) to the same cells. Hint and Help topic. Limits: a calculated item can't be picked in a report
      filter; it shows no details (as Excel); sorting by a value leaves it last.
    - ✅ The rest of PivotTable Options, now Excel's six tabs. Layout & Format: Merge and center cells with labels
      (merged and centred as Excel 365 does, measured: tabular outer items down over their items, totals across the
      label columns, column items across theirs and column totals down; turned off they come apart and lose the
      centring), the compact indent (0–127 characters), report filters Down, Then Over or Over, Then Down with so many
      per column or row (the filter buttons follow), Preserve cell formatting on update. Totals & Filters: Allow
      multiple filters per field (a label filter and a value filter together; each submenu clears its own), Use
      Custom Lists when sorting (days and months first, in calendar order, then the rest A to Z, as Excel). Display:
      contextual tooltips, Classic PivotTable layout (turns tabular form on), item labels with no values, Field List
      Sort A to Z or data source order (the Fields pane follows). Printing: Set print titles (the table's header rows
      and label columns repeat on each page while the sheet has no print titles of its own), with the two other
      print options kept. Data: Save source data with file (off: no records saved; such a file asks for Refresh
      before changes, Excel's message), items to retain (Automatic, None, Max). Alt Text: title and description. All
      read and written as Excel does (the file's own defaults differ from Excel's for indent, multiple filters and
      repeat row labels; Alt Text and hideValuesRow now written into a read file's x14 extension too). Excel-made
      fixtures `pivot-options.xlsx` and `pivot-nodata.xlsx`. Checked in Excel 365: no repair, every option read back,
      and three PivotTables (merged classic, indented with filters over then down, months sorted by custom lists)
      keep every cell, merge, alignment and indent when Excel refreshes them. Hint and Help topic. Limits: no
      contextual tooltips or +/− buttons in print (kept for Excel); Classic layout's drag-into-the-grid isn't there
      (the Fields pane does it); "Show properties in tooltips" and "Show items with no data" (field-level in Cells)
      are not on these tabs.
    - ✅ New PivotTable Style. Excel's dialog: a name ("PivotTable Style 1", 2…), the 25 table elements with Format
      (Font with no font name or size, Border with inside lines and no diagonals, Fill; Clear), Stripe Size for the
      four stripes, the element's formatting in words, a live preview and "Set as default PivotTable style for this
      document". The PivotTable Styles gallery shows the workbook's styles first as Custom; right-click any style for
      Apply, Modify, Duplicate, Delete and Set As Default. Modify renames the PivotTables with it; Delete puts them back
      on Light 16 and drops the default (measured in Excel); new PivotTables take the default. Each change is one undo
      step. A style element can now turn bold, italic, underline and strikethrough on, or off again over the elements
      below (Excel's val="0"), and the report filter labels and values follow their grid. The workbook's own table
      styles (custom table styles from Excel files) are read and drawn too. `<tableStyles>` and its dxfs read and
      written as Excel writes them (kept as they were when unchanged). Excel-made fixture `pivot-custom-style.xlsx`.
      Checked in Excel 365: a Cells-made style opens without repair, as the default, with its stripe size, and Excel
      draws the PivotTable's header and Grand Total as Cells does. Hint and Help topic. Limits: theme colours in a
      style are kept as the colours they show.
    - ✅ New Table Style: the same dialog for tables (Excel's 13 table elements, "Table Style 1", "Set as default
      table style for this document"), from Table Design › Quick Styles and Home › Format as Table. Both galleries
      show the workbook's table styles as Custom, with the right-click menu (Format as Table's has no Apply, as
      Excel's). Insert › Table, Ctrl+T and Quick Analysis take the default table style. Delete puts tables back on
      Medium 2 and drops the default (measured in Excel); style names are shared with PivotTable styles. Excel-made
      fixture `table-custom-style.xlsx`; Excel 365 draws its stripes (two rows, then one) as Cells does, and opens a
      Cells-made table style without repair as the default. Fixed with it: Table Design's gallery drew its thumbnails
      without their fills (the table's own range went into the sample), and a table could be made over a PivotTable
      (now refused with Excel's message).
  - ✅ Get Data › From PDF (asked for 2026-11-26: text PDFs first, scanned ones in the same round), in parts:
    - ✅ PDFs with text. Get Data › From File › From PDF opens the Navigator on the file's tables and pages, named as
      Excel's Pdf.Tables names them (Table001 (Page 1), Page001), each with a preview of what it loads as; Select
      multiple items, Load (a table on a sheet of its own each, columns fitted) or Transform Data (the Query Editor).
      Tables are found by `@cells/io`'s `pdfTables` from the text's positions and ruling lines read with pdf.js
      (`app/src/pdf/`): a grid of lines gives the cells; without lines, rows are the text's lines and columns the
      gaps running down every row (narrow gaps too where the text beside them lines up, as right-set numbers do), a
      value wrapped over lines stays in its row (above or below its middle), a heading set apart goes over its
      column, and a table carried over to the next pages with the same columns is one table ("Table001 (Page 1-3)",
      the repeated heading dropped). Headings are promoted, and types set as Power Query reads text (measured in
      Excel: the region's separators, (5.00) and 5.00- negative, 12% as 0.12, dates in the region's order, ISO and
      month names always); currency signs and spaced digit groups are taken out first by Replace Values steps, and
      another region's dates or numbers get a "Changed Type with Locale" step, so Excel reads them the same. Queries
      are written as Excel's Navigator writes them (Pdf.Tables(File.Contents(path), [Implementation="1.3"]),
      Source{[Id="Table001"]}[Data]); Refresh reads the file again. Recent Sources lists PDFs. Demo fixtures in
      `packages/io/test/fixtures/pdf` (two made by Excel, two by LibreOffice). Checked in Excel 365: Excel finds the
      same cells in the bordered, invoice and ruled-across tables; a Cells workbook with three PDF queries opens
      without repair and Excel's refresh gives the same cells. Hint and Help topic.
      Limits: Excel and Cells each find a PDF's tables their own way, so where they differ (Excel doesn't split a
      wide date column from its description, or join pages) a query refreshed in the other app can give other
      columns; a browser reads the file again only in the session it was picked in; turned and vertical text is left
      out; Excel's other Pdf.Tables options (EnforceBorderLines, StartPage) are kept but not run.
    - ✅ Scanned pages: text recognition on this device, opt-in ("Read scanned PDFs with text recognition" in the
      setup's AI step and File › Options › AI, `apps.cells.ocr`; it covers pictures too). Opening a scanned PDF with it off asks once
      ("This PDF is scanned…"); the Navigator's note on unread scan pages has "Read them with text recognition".
      Tesseract (tesseract.js 7, LSTM, English `best_int` model) runs in a worker in the web view, from files Cells
      ships (`vite.config.ts` copies them to `/ocr/`, nothing fetched from the internet; CSP allows wasm). Pages are
      drawn at 300 dpi, straightened, read as sparse text (a table's rows stay apart), and chunks it wasn't sure of are
      read again on their own as a line, the better reading kept; the words go to the same table finder, a recognized
      chunk's words staying one cell. Cells still unsure are marked in the Navigator's preview ("Check the marked
      cells"). The status bar shows "Recognizing text: page 2 of 5…". Tests read two demo scans (300 dpi, tilted,
      blurred, noisy: `scan-statement.jpg`, `scan-invoice.jpg`) with the real engine; checked in the app in light and
      dark.
      Limits: English only for now; a page takes about a second (more on a slow machine); handwriting and very poor
      scans read badly; the install grows by about 15 MB (engine and model). Checked in Excel 365: Pdf.Tables finds
      no tables in a scanned PDF and its pages come back empty; a Cells workbook with queries from one opens without
      repair with its data, and Excel's Refresh fails (the table isn't there for it) and leaves the data as it was.
    - ✅ Android: Google's ML Kit text recognition with the Latin model bundled (offline, no Play services, nothing
      sent anywhere; ML Kit picks how to run its model on the device, Cells can't choose the NPU). `AndroidOcr.kt`
      is `window.CellsOcr`: the PDF is handed over once and Android's own PDF renderer draws each scanned page at
      300 dpi straight into ML Kit (drawing it in the web view and encoding a JPEG took seconds a page); a page it
      draws blank is drawn by pdf.js and sent as a picture instead. ML Kit doesn't straighten a scan, so the words are
      turned back by its lines' angle before columns are found (0.4° moves a column by a word gap over a page). The
      Android app leaves Tesseract's files out. The prompt says "this device", and a content link's PDF is named by
      the picker's name (it showed "189117"). On the OnePlus Pad: 0.3–0.6 s a page, the demo scans' tables exactly
      right, 1.2 s from picking the file to the Navigator. ML Kit's answers are kept as a fixture
      (`mlkit-scans.json`). Also: Cells' PDF of pictures now says a grey JPEG is grey (Android drew it blank).
    - ✅ Data from Picture (Data › From Picture › Picture From File / Picture From Clipboard): a photo or screenshot
      of a table into cells, read on this device by the same recognizer (Tesseract on the desktop, ML Kit on
      Android, which reads the file itself so a photo's EXIF turn and HEIC work) and the same table finder (the text
      sized to 12 points first, so any picture size gives the same columns). The review shows the picture beside an
      editable grid: the table found, or everything in the picture when there's more (a title, an address); cells
      recognition wasn't sure of are marked, Next to Check goes through them, and Insert Data asks while any are
      unchecked. It goes in at the active cell, as typed in the region (currency and spaced numbers as numbers, a
      leading = kept as text), a date column written in another order than the region's read in that order
      (13/01/2026 on a US computer), columns widened where needed, one undo step, asking before replacing data.
      Turning recognition on is asked once (the same switch, now "scanned PDFs and pictures"). Hint "New: Data from
      a picture" and a Help topic. Checked in the browser (light, dark, phone, a sideways photo with its EXIF turn)
      and on the OnePlus Pad: a screenshot and the invoice scan read right in about 1.2 s.
      Limits: Picture From Clipboard is off on Android (no picture clipboard there); no camera capture yet;
      English / Latin text only; no Excel file part (it only puts in values).
  - ✅ Searchable PDF text: File › Export › Create PDF lays each page's text invisibly over its picture, where it
    was drawn (as a scanner's searchable PDF), so the PDF can be searched and its text selected and copied, while
    the page still looks exactly as on screen. The text is recorded as the page draws it (`print/text-layer.ts`
    wraps the canvas's fillText: transform, alignment, baseline, turned text) and set in a 512-byte shapeless
    TrueType font Cells writes itself (`pdf-font.ts`, as Tesseract's GlyphLessFont) with a ToUnicode map, so any
    character copies right (accents, CJK, emoji). Checked with pdftotext, pdffonts and Ghostscript (the page renders
    the same, the words' boxes sit on the drawn words, wrapped, right-aligned and turned text included) and pdf.js:
    Cells' own Get Data › From PDF reads an exported table back cell for cell.
    Limits: text inside charts, shapes and pictures isn't in the layer (they're drawn as pictures); a cell's text
    is there in full where the picture clips it (as Office's PDFs keep clipped text); diagonal text comes out in
    pieces in some readers; the system Print dialog's "Save as PDF" still makes pictures only.
  - ✅ Notes and comments that move with their cells, as measured in Excel 365: Insert Cells (shift down or right)
    and Delete Cells (shift up or left) move them along their rows or columns, a deleted cell's note going with it;
    cut and paste moves them, on the sheet or to another, replacing the destination's; copy and paste copies them,
    tiled with the cells, also replacing the destination's (a cell without a note pasted over one takes it away).
    Sort already moved them with their rows; Fill and text pasted from other apps leave them, as in Excel. Every
    case undoes in one step. Checked in the app: a note's red corner follows its cell.
  - ✅ Name Box lists: the Name Box takes a list of ranges (`B1,D3:E4`, or with semicolons) and selects them all,
    the last one's first cell active, as Excel goes to one (a list across sheets is refused); a name made from a
    selection of several ranges can now be gone to, and the Name Box shows it when that selection is made again. A
    table's name goes to its data rows (without the header and total rows, as in Excel). The Name Box's list holds
    the names that point at cells and the tables, A to Z (constants such as `=0.2` left out), built only while open.
  - ✅ Live Flash Fill (Excel's "Automatically Flash Fill", on by default, File › Options › Advanced › Editing
    options; `apps.cells.autoFlashFill`): typing in an empty cell under an example or more, beside the data, shows
    the rest of the column in grey once the examples' pattern gives a value starting with what's typed (with a
    header row over the column or without it); the grey cells cover text running over from the left, as filled
    cells do. Enter takes it: the typed cell gets its whole value and the column the rest, one undo step named
    "Flash Fill"; Tab, Esc, an arrow or a click leave it. No offer for formulas, cells that already hold a value, or
    the first entry in a column. The pattern is learned once per cell and only run as the entry grows. Hint "New:
    Flash Fill as you type" (a new id, so it shows to those who dismissed the old one).
    Limits: no Flash Fill Options button after it (Undo / Accept / Select changed cells); the offer comes from the
    cheapest pattern the examples fit (Excel weighs several); Excel's exact rule for when the grey appears isn't
    published, so it may show sooner or later than Excel's.
  - ✅ .xlsb and .ods. **Excel Binary Workbooks** open as a copy that saves as .xlsx (.xlsm when it has macros), as
    .xls do: the BIFF12 records are written out as the SpreadsheetML they stand for and read by the .xlsx reader, so
    values, formulas (shared, array and dynamic, tables' structured references, 3-D and names), styles, sizes,
    merges, notes, links, names, panes, tables, hidden sheets and page setup come along; checked against Excel's
    .xlsx of the same workbooks. **OpenDocument Spreadsheets** open as themselves and save back (Save As ›
    OpenDocument Spreadsheet; Save asks "Some features … might be lost" once per workbook, as Excel does): values,
    OpenFormula formulas (`;`, `|`, `~`, `!`, `COM.MICROSOFT.`, LET's `_xlpm.` names, renamed functions), cell styles
    and number formats, column widths and row heights, hidden rows/columns/sheets, tab colours, merges, notes,
    links, named ranges and expressions, print ranges, frozen panes, zoom and gridlines; everything recalculates on
    opening, as Excel does. Reading checked against Excel's and LibreOffice's .ods of one workbook; what Cells
    writes opened in LibreOffice and in Excel and compared cell by cell. Open/Save types, desktop file
    associations and Android's Open with list both.
    Limits: .xlsb rich-text runs are dropped, and conditional formats, data validation, pivot tables, external links
    and connections don't open (a warning says so). .ods charts, pictures, shapes, conditional formats, validation
    and DataPilot tables don't open (a warning says so) and aren't saved; tables save as plain ranges with their
    structured references written as the cells they stand for (as Excel does); a spilled-range reference (A1#)
    saves as its value. LibreOffice sizes columns and fitted rows with its own fonts, so they can differ by a few
    pixels; an intersection that spills (`=H1:H3 H2:H5`) is #VALUE! in LibreOffice, #SPILL! in Excel.
- ⏸ Out by design: CUBE*, RTD, CALL, REGISTER.ID, EUROCONVERT, COPILOT, TRANSLATE, DETECTLANGUAGE, FIELDVALUE,
  STOCKHISTORY (online services, add-ins or Data Model).

**Known limits:** distributions, tests and fits computed through special functions may differ from Excel in the
16th significant digit (both are within a few units of the exact value, checked against 50-digit arithmetic);
`tests/excel-golden/tolerances.json` lists them with how far. Excel's BESSELJ is itself only good to about 2e-5
for x ≥ 8; Cells is closer to the true value there, so the two differ in the 5th digit.
Excel's YIELD and ODDFYIELD stop iterating early (agree to 9 digits); a few Excel bond results fit none of its
documented conventions (two actual/actual COUPDAYS periods at month ends, ODDF* long first periods on actual/360 and
actual/365) and are listed in `tests/excel-golden/known-differences.json`.
GROUPBY/PIVOTBY: one function at a time (an HSTACK of functions isn't supported yet); field_relationship 1 drops
subtotals but sorts as a hierarchy; generated headers ("Row Field 1", "Value") and an all-blank TRIMRANGE (#REF!)
follow our reading of the documentation, not a check in Excel. Trim references (A1:.E10) aren't parsed yet.
REGEX* use JavaScript regular expressions rather than Excel's PCRE2 (no inline (?i), possessive quantifiers or
\K); FILTERXML covers XPath 1.0's paths, predicates and common functions, not namespaces. ASC, DBCS and the *B
functions behave as in a language without double-byte characters (as Excel does there); PHONETIC gives #N/A, as no
cell carries furigana. CELL("filename") and INFO("directory") are empty, CELL without a reference uses the formula's
own cell (Excel: the last changed cell), INFO("recalc") always says Automatic.

---

### Phase 9 — Trust, printing, shapes and the rest of the gaps (0.9.0) 🔄

Chosen 2026-11-26 from the shortlist (the .ods/.xlsb round two stays for later). In this order, one commit per
part; the verify sweep runs alongside, area by area, as each part touches it.

- ✅ 1. Quiet fixes: links move with cut and paste and with Insert/Delete Cells, and a pasted cell brings its link;
  an Advanced Filter in place is saved as Excel saves it (_FilterDatabase, Criteria) so Clear and Reapply work
  after reopening; "Select locked cells" off is enforced on a protected sheet; rows refit when a column width
  changes, and edits over 20,000 rows refit too; the wheel scrolls by rows, as in Excel (a touchpad stays smooth).
- ⬜ 2. Verify sweep: the open items in `docs/verify-list.md` checked in Excel 365 and on the OnePlus Pad / Pixel,
  fixed where they differ, ticked off as they pass.
- ✅ 3. Printing and page layout: ✅ different first page and odd/even headers and footers (the Options group, the
  Page Setup dialog's options and Custom Header tabs; Excel picks them by the page's place, not its number); ✅
  Picture / Format Picture in headers and footers (Header & Footer › Picture puts &[Picture] at the caret, asking
  before replacing a section's picture; Format Picture's Size and Picture tabs: size, scale, aspect lock, crop,
  colour, brightness, contrast, Reset; drawn in line with the section's text in Page Layout and printing; saved as
  Excel does in a legacyDrawingHF VML drawing, one image part per picture; checked in Excel 365, which reads every
  setting back); ⏸ print titles repeated on each page in Page Layout view (left out: Excel
  shows them only when printing, as Cells does); ✅ black-and-white printing applied (no fills, text and borders
  black, objects in greys, as Excel's PDF shows); ✅ printing and PDF of chart sheets (one page, the chart filling it inside the margins).
- 🔄 4. Shapes and text boxes: ✅ formatted text in shapes (typed in place in its own fonts, sizes and colours; Home's
  Font and Alignment commands and Ctrl+B/I/U format the selected text while typing, or all of a selected shape's
  text; typed text takes the look it's typed into; equations and turned text keep the plain editor); ✅ WordArt
  Styles gallery on Shape Format (with Clear WordArt; Shape Styles was already there); ✅ adjustment handles (the yellow dots on 21 shapes: corners, arrow heads and shafts, callout points;
  a dashed outline while dragging, one undo step); ⬜ the Format Shape pane, Edit Points and freeform drawing,
  connectors that stay glued to shapes.
- ⬜ 6. Proofing: more spelling languages as optional downloads (setup and Settings, the language picker in the
  Spelling dialog), File › Options › Proofing (ignore UPPERCASE, words with numbers, addresses), AutoCorrect,
  checking text in shapes, comments and headers/footers.
- ⬜ 7. Themes: Customize Colors and Customize Fonts, Save Current Theme, Reset to Theme from Template; conditional
  formats and chart colours from files following the theme.
- ⬜ 8. Several workbooks at once: each open workbook in its own window on the desktop (and a switcher on phones and
  tablets), with every window's view saved (Excel's workbookView per window).

To decide when we get there: where the spelling dictionaries are downloaded from (6); how several workbooks look
on Android (8).

## 4. Verifying against Excel (all phases)

- **Mechanism:** local, git-ignored tooling like Inkspire's Word renderer: copy files into the Office VM through the
  QEMU guest agent, drive a hidden Excel instance with PowerShell/COM, bring results back. Nothing on the host desktop
  is captured; screenshots of Excel come from the VM itself, ours from Xvfb.
- **Tools to write (Phase 1–2):** `excel-run.py` with modes `eval` (write inputs + formulas to a workbook, recalc,
  read back values/types/errors as JSON), `format` (`TEXT(value, code)` table), `roundtrip` (open our file, report
  repairs, dump every cell's value, formula and format, re-save), `render` (PDF / PNG of a sheet range).
- **Golden data, committed:** only the JSON expectations and synthetic fixtures (no VM details, no real files);
  each expectation records the Excel version it came from.
- **When Excel and we differ:** fix, or record it in `docs/excel-differences.md` with the reason. Nothing is silently
  skipped.

---

## 5. Working conventions (copied from Inkspire)

- `docs/changes.md`: a dated entry per fix/feature (what was wrong, what changed, how it was checked, what is open).
- `docs/next-todo.md`: the active phase's checklist; later phases are not started until requested.
- `docs/verify-list.md`: items built but not yet checked against Excel or on a device.
- Tests are written with the feature; each phase closes with lint, typecheck, Vitest, cargo test and the Excel
  checks green, and a short phase report (what was built, numbers measured, known differences).
- Privacy: fixtures, screenshots and docs use made-up names and data only; commits use the GitHub noreply identity;
  diffs are grepped for personal names, emails and real paths before each commit.

---

## 6. Risks

| Risk                                                              | Mitigation                                                                                              |
| ----------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| Excel's formula semantics are vast and full of quirks             | Conformance suite generated from Excel itself; functions shipped in waves; unknown functions round-trip |
| Canvas text rendering differs across WebKitGTK/WebView2/WKWebView | Measure-text cache per font; screenshot tests per platform in CI                                        |
| 100k+ rows in a webview                                           | Worker engine, viewport-only patches, tile cache, budgets enforced by benchmarks                        |
| `.xlsx` fidelity (pivots, VBA, slicers we do not implement)       | Preserve unknown parts verbatim; corpus + Excel open checks                                             |
| Fonts: Aptos/Calibri absent on Linux → different column fits      | Metric-compatible fallbacks (Carlito for Calibri); note differences in comparisons                      |
| Phone UX for a dense grid                                         | Separate PhoneShell, tested on a real device, not a squeezed desktop                                    |
| Scope creep from "match Excel"                                    | Each phase has an explicit list; anything else goes to `next-todo.md` for a later decision              |

---

## 7. Decisions log

All open questions were answered on 2026-11-14:

- Look: Office-like style and Inkspire's tokens (not FileSort); own app icon.
- Name and repo: **Cells**, its own GitHub repo, `main` branch only (Android included).
- Inkspire integration: match Inkspire's local-only setup; no accounts, cloud or sync.
- Changes needed in Inkspire: deferred, tracked in `docs/inkspire-todo.md`, done later in the Inkspire project.
- Reference VM: the existing Office VM; Excel is installed there.
