# Linked ranges and Edit Data (Cells ↔ Inkspire)

PLAN.md Phase 5, item 5; Inkspire's side is `docs/inkspire-todo.md` I5 and I7. Both work through files, so neither
app has to start or find the other: whichever runs reads what the other left.

## Live linked tables

The folder is `links/` beside the shared settings (`<config dir>/inkspire/links/`, `~/.config/inkspire/links/` on
Linux).

**Inkspire writes `wanted.json`**: the ranges the documents it has open link to (its `LINK Excel.Sheet.12` fields),
rewritten when a document opens, closes or gets a new link:

```json
{ "version": 1, "links": [{ "file": "/home/alex/Q1 Budget.xlsx", "range": "Sheet1!R1C1:R10C4" }] }
```

`range` is as the LINK field has it (R1C1, or A1, or a defined name).

**Cells writes `<name>.json`** for the workbook it shows, when Inkspire wants any of its ranges: on opening it, and
a moment (0.4 s) after anything changes (an edit, a recalculation, a save) or Cells gets the focus. A snapshot is
written only when it differs from the last one. `<name>` is the FNV-1a 64-bit hash of the workbook path's UTF-8 bytes,
as 16 lowercase hex digits (`""` → `cbf29ce484222325`, `"a"` → `af63dc4c8601ec8c`).

```json
{
  "app": "cells",
  "version": 1,
  "file": "/home/alex/Q1 Budget.xlsx",
  "updated": "2026-11-15T10:30:00.000Z",
  "saved": false,
  "ranges": {
    "Sheet1!R1C1:R10C4": {
      "sheet": "Sheet1",
      "range": "A1:D10",
      "text": "Item\tCost\r\n…",
      "html": "<table …>…</table>",
      "values": [
        ["Item", "Cost"],
        ["Rent", 1200]
      ]
    },
    "Totals": { "error": "#REF!" }
  }
}
```

- `html` is the same formatted table a copy gives (`docs/inkspire-clipboard.md`); `values` holds numbers, text,
  true/false, error codes (`"#DIV/0!"`) or null.
- `saved: false` means the workbook has changes not yet saved to `file` (Word shows such changes too, while Excel has
  the workbook open).
- A range that doesn't exist is `{ "error": "#REF!" }`; one over 20,000 cells is `{ "error": "too big" }`.
- Files are written whole and renamed into place, so a reader never sees half of one.

**Inkspire's rule:** watch the folder; for a link, use the snapshot when its `updated` is later than the workbook
file's modification time, and otherwise read the workbook file itself (Cells closed, or the file changed since).

## Edit Data

For a chart whose data lives in the document (the embedded workbook Word and Inkspire write), Inkspire:

1. writes the embedded workbook to its edit folder, `<cache dir>/inkspire/edit/` (`~/.cache/inkspire/edit/` on
   Linux), as an `.xlsx`;
2. opens `cells://edit?file=<that path>&for=<document name>`;
3. watches the file and reads the chart's data back whenever it changes.

Cells shows the workbook as "Chart in Report.docx" and saves each change back to the file 0.8 s later (not while a
cell is being typed in). Closing it doesn't ask about saving, and it isn't added to Recent. Save As makes it an
ordinary workbook again. Cells saves back only to files inside Inkspire's edit folder; for any other file the link
just opens it, as `cells://open` does.

Code: `app/src/links/live.ts`, `FileCommands.openEmbedded` in `app/src/file/commands.ts`, `src-tauri/src/lib.rs`
(`read_link_wants`, `write_link_snapshot`, `is_edit_copy`).
