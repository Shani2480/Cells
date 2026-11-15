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
- [ ] I4. **Paste from Cells.** Handle the Inkspire range clipboard type (full styles, chart spec, source workbook and
      range); fall back to the HTML/TSV that Cells also puts on the clipboard.
- [ ] I5. **Linked tables.** `LINK Excel.Sheet.12 "file.xlsx" "Sheet1!R1C1:R10C4" \a \f 4 \h` fields: refresh on open,
      "Update Link" on demand, live refresh when Cells announces a change; saved so Word can update the link too.
- [ ] I6. **Linked charts.** Chart parts with `c:externalData` pointing to an external workbook, read and written.
- [ ] I7. **"Edit Data in Cells".** For an embedded chart, open its embedded `.xlsx` in Cells instead of
      the docked data-sheet dialog, and write the result back into the document.
- [ ] I8. **Shared code** (after Cells Phase 4): adopt the shared OPC/zip, chart spec/XML, ribbon and token packages,
      and replace the 4×4 Spreadsheet block (`editor/sheet-block.ts`) and arithmetic formula fields
      (`editor/formula.ts`) with the Cells formula engine.
