# Changes

Newest first. Each entry: what changed, how it was checked, what is still open.

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
