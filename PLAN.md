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
Preview, sparklines, the chart Format tab and `.ods`. PDF pages are images for now. What's left to check
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

### Phase 7 — The commands marked "not built" (0.7.0) 🔄

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
- ⬜ Next: the Help tab.
- ⬜ Then: Inkspire's side of the link (`docs/inkspire-todo.md`, I1–I9), done in the Inkspire project.
- ⏸ Later: the lists under "Later" in `docs/next-todo.md` (among them: a More Sort Options dialog for PivotTables,
  Field Headers, Show Report Filter Pages, calculated items, new PivotTable styles).

**Known limits:** sorting a PivotTable by a value is set only by the AI or by Excel's files (no dialog yet).

---

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
