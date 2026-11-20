# AI tools

Home › AI Tools, in the place Excel has Copilot (PLAN.md Phase 5, decided 2026-11-15: Cells uses Inkspire's AI
providers and keys, opt-in). Off until the person turns it on, in the first-run setup's last step or File › Options ›
AI.

## The tools

| Tool                | Sends                                              | Gives                                             |
| ------------------- | -------------------------------------------------- | ------------------------------------------------- |
| Formula from Words… | the description, the cells around the active cell  | a formula, checked to parse, inserted on Insert   |
| Explain Formula     | the active cell's formula and the cells around it  | a short explanation (Copy)                        |
| Suggest a Chart     | the selection (or the data around the active cell) | one of nine chart types and a title; Insert Chart |
| Summarize Data      | the selection (or the data around the active cell) | three to five lines (Copy)                        |
| Describe PivotTable | the data's column names, kinds and a few values    | a PivotTable layout, previewed; OK or Apply       |

Cells are sent as `A1: Item | B1: Cost` lines, at most 60 rows and 26 columns, with formulas as written and the value
they show. Every dialog says what it sends and where ("…to Anthropic, outside this computer" or "…on this computer"),
and nothing goes into the workbook until the person chooses Insert or Insert Chart (one Undo step). A reply that isn't
a formula Cells can read, or a chart Cells draws, is shown as a problem instead.

"Describe the PivotTable you want" lives in Recommended PivotTables and the PivotTable Fields pane, and only shows
while the AI tools are on and set up. It never sends rows: a text column is described by up to six of its values, a
number or date column by its smallest and largest. The model only picks the layout (`app/src/ai/pivot-prompt.ts`,
`DescribePivot.tsx`, `DescribePivotDialog.tsx`); Cells' PivotTable engine works out every number.

Code: `app/src/ai/prompts.ts` (requests and reading replies), `app/src/ai/AiDialog.tsx`, `app/src/ai/client.ts`
(where requests go), `app/src/settings/AiFields.tsx` (Options and setup), `src-tauri/src/ai.rs`.

## Providers and keys (shared with Inkspire)

- **Choice:** `ai.provider`, `ai.models.<provider>` and `ai.addresses.<provider>` in the shared settings file
  (`docs/shared-settings.md`), with Inkspire's names: `builtin`, `ollama`, `lmstudio`, `compatible`, `openai`,
  `anthropic`. Whether the tools are on is Cells' own (`apps.cells.ai`).
- **Keys:** the system keyring, service `com.aiword.app.ai`, account `openai`, `anthropic` or `compatible`, as
  Inkspire's `secrets.rs`. Without a keyring, Inkspire's owner-only file `<config dir>/com.aiword.app/ai-keys.json`.
  `OPENAI_API_KEY` and `ANTHROPIC_API_KEY` work too. Keys stay in the Rust process; the webview only learns whether
  one is saved.
- **Requests** go from Rust (`ai_generate`, streamed as `ai-chunk` events, `ai_cancel` stops one). Ollama only on this
  computer; an OpenAI-compatible endpoint (LM Studio) over https, or http on this computer or the local network.
- **Not yet:** Inkspire's built-in engine (it downloads and runs a llama.cpp server; `local_ai.rs`). Cells says so and
  points to Ollama or LM Studio for on-device AI. Sharing Inkspire's engine and model files is the next step.
- The browser build has no AI (it can't keep keys or reach local models). `?ai=demo` in the dev build gives canned
  answers, for trying the dialogs.
