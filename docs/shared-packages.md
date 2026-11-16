# Shared packages (Cells and Inkspire)

PLAN.md Phase 5 item 6 left the choice of what to share until both apps had the same code. Decided 2026-11-16:
share what is really the same code, from this repo, as plain TypeScript packages that Inkspire builds with its own
Vite. Inkspire's side is item I8 in `inkspire-todo.md`; nothing in Inkspire changes from here.

## What is shared

| Package        | What it is                                                                                                                              | Replaces in Inkspire                                                                                                                |
| -------------- | --------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `@cells/ooxml` | Zip reading and writing, XML reading and escaping, and the OPC layer (content types, relationships, part order) for any Office package  | `src/document/docx/xml.ts`, `zip.ts`, `opc.ts` (with `WORD_PACKAGE`; see below)                                                     |
| `@cells/chart` | The chart model, its SVG renderer, DrawingML and chartex parts, the embedded data workbook                                              | `src/editor/chart-spec.ts`, `chart-expand.ts`, `chart-map.ts`, `chart-sheet.ts`, `src/document/docx/chart-xml.ts`, `chartex-xml.ts` |
| `@cells/core`  | The spreadsheet, including the formula engine. `calculateGrid` and `calculateFormula` are its entry points for grids outside a workbook | the arithmetic in `src/editor/sheet-block.ts` (`resolveValue`) and `src/editor/formula.ts` (`evaluateFormula`)                      |

`@cells/io` (reading and writing `.xlsx` and `.csv`) is Cells' own; Inkspire doesn't need it.

### `@cells/ooxml`

The same three files Cells first took from Inkspire. What differed between a workbook and a Word document now comes
in as a `PackageFormat`: `SPREADSHEET_PACKAGE` (Cells) and `WORD_PACKAGE` (Inkspire: `.docx/.docm/.dotx/.dotm`,
Word's part order and passthrough parts, the messages Inkspire shows). Inkspire keeps its function names with a
three-line wrapper, as Cells does in `packages/io/src/opc.ts`:

```ts
import { openPackage as open, writePackage as write, WORD_PACKAGE, type OpcPackage } from "@cells/ooxml"
export const openPackage = (bytes: Uint8Array, fileName = "document.docx") => open(bytes, fileName, WORD_PACKAGE)
export const writePackage = (pkg: OpcPackage) => write(pkg, WORD_PACKAGE)
```

`isPassthroughPart(part)`, `allowsMacros(name)` and `mainContentTypeFor(name)` take the format as a second argument;
`WORD_PACKAGE_EXTENSIONS` is `packageExtensions(WORD_PACKAGE)`. `CONTENT_TYPE_MACRO` and `CONTENT_TYPE_TEMPLATE` are
now `CONTENT_TYPE_DOCUMENT_MACRO` and `CONTENT_TYPE_DOCUMENT_TEMPLATE` (the workbook ones carry `WORKBOOK`).

### `@cells/chart`

Copied from Inkspire for Cells' Phase 4. Since then it differs only in its imports, a `clone` helper for
`structuredClone`, and an optional `refs` argument to `chartSpaceXml` (the cell references a chart in a workbook
writes; a document's chart leaves it out). Inkspire's `THEME_ACCENTS` and `shade` stay in `shape-style.ts` for
shapes; the package has its own copy in `theme.ts`.

### The formula engine

```ts
import { calculateGrid, calculateFormula, gridColumns } from "@cells/core"

calculateGrid([
  ["Rent", "1,200"],
  ["Food", "=B1/4"],
])
// → [[{ value: "Rent", text: "Rent" }, { value: 1200, text: "1,200" }], [{ value: "Food", … }, { value: 300, text: "300" }]]
calculateFormula("=ROUND(PI(), 2)") // → { ok: true, value: 3.14, text: "3.14" }
calculateFormula("=2+*") // → { ok: false, error: "Invalid formula" }
calculateFormula("=1/0") // → { ok: false, error: "#DIV/0!" }
```

Cells are read as Cells reads typed text (numbers with separators, percentages, dates, `'` for text) and formulas
have Excel's functions, so the Spreadsheet block can grow past 4×4 without new code. Both take
`{ regional: regionalFrom(sharedSettings) }` to read and show numbers in the person's formats. Word's own field
messages are Inkspire's to choose: `"Invalid formula"` is Word's `!Syntax Error`, `"#DIV/0!"` its `!Zero Divide`.
Word's `=SUM(ABOVE)` table references aren't Excel syntax; Inkspire would turn them into a range first.

## What isn't shared, and why

- **Ribbon components.** Inkspire builds its ribbon as one component per tab; Cells builds it from a command registry
  (`app/src/ribbon/registry.ts`) that the classic ribbon, keytips, the QAT, Tell Me and Customize Ribbon all read.
  The two share no component code, so a package would mean rewriting one app's ribbon. Both follow `office-style-ui`,
  which is where what they share (sizes, behaviour) is written down.
- **Tokens.** 19 CSS variable names are the same in both apps (Cells has about 80, Inkspire about 190), and each has
  its own theme colour. A shared token file would first need the names lined up; until then each app keeps its own,
  with the names Cells took from Inkspire noted in `app/src/styles/index.css`.

## How Inkspire uses them

1. Add this repo as a git submodule at `vendor/cells`, pinned to a commit (updating it is a deliberate step).
2. Add the three packages to Inkspire's npm workspaces and as dependencies:
   ```json
   "workspaces": ["vendor/cells/packages/ooxml", "vendor/cells/packages/core", "vendor/cells/packages/chart"],
   "dependencies": { "@cells/ooxml": "*", "@cells/core": "*", "@cells/chart": "*" }
   ```
   The packages export TypeScript source, which Inkspire's Vite and `tsc` build like its own files.
3. Replace the files in the table above with imports, keeping Inkspire's own names through small wrappers where
   that saves touching many call sites.

Checked 2026-11-16: the three packages compile under Inkspire's `tsconfig.json` (ES2020 + DOM, `strict`,
`noUnusedLocals`, bundler resolution) with no errors, using the `@types/d3-geo` and `@types/topojson-client` Inkspire
already has.

## Rules for the shared packages

- Pure TypeScript: no DOM, no React, no Tauri, no `@/` paths (`scripts/check-boundaries.mjs` checks the imports).
- ES2020 library only (their `tsconfig.json` says so), because Inkspire builds with it.
- A change that alters an export's behaviour or signature notes it here and in `inkspire-todo.md`, so the next
  submodule update in Inkspire knows what to adjust.
