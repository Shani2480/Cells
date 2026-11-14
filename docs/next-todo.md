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

## Phase 2 — done, waiting for the go-ahead for Phase 3

- [x] Lexer and parser (A1, R1C1, 3-D, arrays, spill `#`, `@`), printing
- [x] Values and coercion
- [x] Dependency graph, dirty marking, volatiles, circular references, manual mode, F9
- [x] Dynamic arrays and spills
- [x] Structural edits rewrite formulas (insert/delete, move, copy, fill)
- [x] 262 functions, Excel conformance suite (`tests/excel-golden/`)
- [x] Formula bar and in-cell editing: colours, Point mode, F4, autocomplete, tooltips, Ctrl+`, Ctrl+Shift+Enter
- [x] Engine in a Web Worker
- [ ] Chain benchmark under 200 ms on a quiet machine (measured 150–200 ms quiet, more under load)

## Later (from Phase 2)

- Version and What's new page in File → Account (Phase 3 backstage), shown once after an update; the notes are
  in `app/src/about/release-notes.ts`

- LAMBDA and its helpers (MAP, REDUCE, BYROW, …)
- R1C1 display option, defined names (Name Manager), iterative calculation
- Calculation options in the ribbon (Formulas tab, Phase 3)
- ROMAN's concise forms; legacy `{=…}` array formulas
- RATE, IRR and XIRR stopping where Excel's solver does
