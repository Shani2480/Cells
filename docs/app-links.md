# Links between Cells and Inkspire

Each app registers a URL scheme and runs one instance at a time (Tauri's `deep-link` and `single-instance` plugins).
A link opened while the app runs goes to the running window, which comes to the front; otherwise the app starts
with it. PLAN.md Phase 5, item 3; Inkspire's side is `docs/inkspire-todo.md` I3.

## Cells

```
cells://open?file=<full path>&range=<range>
```

- `file`: the workbook's full path, percent-encoded (`%2Fhome%2Falex%2FQ1%20Budget.xlsx`). Only workbook types Cells
  opens (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.csv`, `.tsv`, `.txt`); anything else is refused.
- `range` (optional): what the Name Box takes (`Sheet1!A1:D10`, `'My Sheet'!B2`, a defined name), or the R1C1 form
  of Word's `LINK` field (`Sheet1!R1C1:R10C4`).

Cells opens the file (asking to save the workbook it replaces, as File › Open does) and selects the range. A link to
the workbook already shown only selects the range. A workbook path on the command line (`cells Budget.xlsx`, a file
opened from the file manager) opens the same way.

A link can only open and show a file: it never saves, changes or runs anything. The one exception is Edit Data:

```
cells://edit?file=<full path>&for=<document name>
```

opens the workbook of a chart in an Inkspire document and saves each change back to it, but only for a file inside
Inkspire's edit folder (`<cache dir>/inkspire/edit/`); any other file just opens. See `docs/linked-ranges.md`.

Code: `app/src/links/link.ts` (reading and writing links), `src-tauri/src/lib.rs` (`deliver`, `take_pending_opens`).

## Inkspire (to build there)

```
inkspire://open?file=<full path>
```

The same rules: percent-encoded full path, documents only, one instance.
