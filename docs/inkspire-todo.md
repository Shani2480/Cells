# Inkspire changes needed for Cells integration (deferred)

Work for the **Inkspire** project, not this repo. Collected here so it can be picked up in Inkspire later
(see PLAN.md, Phase 5). Tick an item (`[x]`, date) when it lands in Inkspire.

- [ ] I1. **Shared settings file.** Move user name, regional formats and theme out of webview `localStorage`
      (`aiword.user-settings`, `aiword-color-mode`) into the shared file described in `docs/shared-settings.md`
      (Cells reads and writes it since 0.5.0). On first start, copy the existing values in where the file has none;
      keep the first-run setup's "done" under `apps.inkspire.setupDone`. Only write the values that changed, and
      read the file again on window focus.
- [ ] I2. **Shared recent-files list** (optional): both apps' File views can show each other's recent files.
- [ ] I3. **URL scheme + single instance.** Register `inkspire://open?file=…` with the Tauri `deep-link` and
      `single-instance` plugins, so Cells can open a document in the running Inkspire. Cells' side and the link
      format are in `docs/app-links.md`; "Edit in Cells" opens `cells://open?file=…&range=…`.
- [ ] I4. **Paste from Cells.** What Cells copies is in `docs/inkspire-clipboard.md`. A range already pastes as a
      formatted table and a chart as an editable Inkspire chart (Cells writes Inkspire's own figure markup), so nothing
      is needed for a static paste. To do: keep `data-cells` (file, sheet, range, link) for Paste Special › Paste Link
      (I5) and an "Edit in Cells" action; keep `data-cells-chart` on a pasted chart for "Edit Data in Cells" (I7).
      Check that a pasted table keeps the column widths and borders.
- [ ] I5. **Linked tables.** `LINK Excel.Sheet.12 "file.xlsx" "Sheet1!R1C1:R10C4" \a \f 4 \h` fields: refresh on open,
      "Update Link" on demand, live refresh from Cells; saved so Word can update the link too. Cells' side is done
      (`docs/linked-ranges.md`): write the open documents' links to `links/wanted.json`, watch the folder, and take a
      snapshot when it is newer than the workbook file (otherwise read the file).
- [ ] I6. **Linked charts.** Chart parts with `c:externalData` pointing to an external workbook, read and written.
- [ ] I7. **"Edit Data in Cells".** For an embedded chart, open its embedded `.xlsx` in Cells instead of
      the docked data-sheet dialog, and write the result back into the document. Cells' side is done
      (`docs/linked-ranges.md`): write the workbook to `<cache dir>/inkspire/edit/`, open
      `cells://edit?file=…&for=<document name>`, watch the file and re-read the chart's data when it changes.
- [ ] I9. **AI settings in the shared file.** Move `aiword.ai-settings` (provider, models, addresses) from
      `localStorage` to `ai.provider`, `ai.models` and `ai.addresses` in the shared settings file (`docs/ai.md`), once,
      keeping what is there; then read and merge them as the other settings. Keys are already shared (same keyring
      service). Later: let Cells use the built-in engine (share `local_ai`'s engine and model folder, or run one
      server both apps talk to).
- [ ] I10. **Shared dictionary words.** Keep words added with Add to Dictionary in `proofing.customWords` of the shared
      settings file (`docs/shared-settings.md`), adding to the list there instead of only Inkspire's own store, and
      load them into the speller, so a word added in either app is known in both.
- [ ] I8. **Shared code** (`docs/shared-packages.md`): add Cells as a submodule at `vendor/cells` and its
      `ooxml`, `core` and `chart` packages to the npm workspaces. Then: - replace `document/docx/xml.ts`, `zip.ts` and `opc.ts` with `@cells/ooxml` (`WORD_PACKAGE`, through a small
      wrapper that keeps `openPackage`/`writePackage`); - replace `editor/chart-spec.ts`, `chart-expand.ts`, `chart-map.ts`, `chart-sheet.ts` and
      `document/docx/chart-xml.ts`, `chartex-xml.ts` with `@cells/chart`; - calculate the Spreadsheet block (`editor/sheet-block.ts`) with `calculateGrid` and formula fields
      (`editor/formula.ts`) with `calculateFormula`, showing Word's `!Syntax Error` / `!Zero Divide`.
      Ribbon components and tokens stay separate (why: `docs/shared-packages.md`).
