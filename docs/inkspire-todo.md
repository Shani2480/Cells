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
- [ ] I11. **Equations from Cells.** Cells has an Office Math engine in its shared packages: the linear format
      (UnicodeMath) parser and printer and the math tree (`@cells/core` `math/`), an OMML reader and writer
      (`@cells/io` `xlsx/omml.ts`, DrawingML run properties; Word's `w:rPr` needs a variant), and an SVG layout
      (`app/src/drawing/math.ts`, to move into a package before sharing). Word's `m:oMathPara`/`m:oMath` in paragraphs
      could use them for display and editing.
- [ ] I12. **3D models from Cells.** Word keeps 3D models the same way (`am3d:model3d` in a graphic frame, the .glb by a
      `model3d` relationship, a picture fallback). Cells' pieces: the view and its XML (`@cells/core` `model3d.ts`,
      `@cells/io` `xlsx/model3d.ts`: the frame needs `wp:inline`/`wp:anchor` instead of `xdr`), and the renderer and
      file import (`app/src/drawing/model3d.ts`, three.js; to move into a package before sharing).
- [ ] I13. **SmartArt from Cells.** Word keeps SmartArt in the same four diagram parts (and Office's drawing). Cells'
      pieces: the layouts and colours written in Office's layout language, the data part from the text tree and the
      tree editing (`@cells/core` `smartart.ts`), the parts read and written (`@cells/io` `xlsx/smartart.ts`: the frame
      needs `wp:inline`/`wp:anchor` instead of `xdr`), the layout engine that draws them as Excel does
      (`app/src/drawing/smartart.ts`, to move into a package before sharing), and the dialog and Text Pane.
- [ ] I14. **Chart part formats kept on reopen.** `@cells/chart` `xml.ts` now writes each part's text size and bold,
      line widths and the plot area's border, and `readChartSpace` reads them back (`readPartFormats`), keeping only
      what differs from the chart style so Word's and Excel's theme colours don't become overrides. Inkspire's copy
      (`editor/chart-spec.ts` and its XML) reads back only fills, so a title colour or a line weight set in Word's
      Format pane is lost on reopening; the same change fixes it. Cells' Recommended Charts rules
      (`app/src/sheet/chart-recommend.ts`) could also feed Word's Insert Chart dialog.
- [ ] I15. **Spelling languages and options, shared.** Cells now keeps the spelling language in `proofing.language`
      and Office's three proofing options in `proofing.ignoreUppercase` / `ignoreNumbers` / `ignoreAddresses`
      (docs/shared-settings.md). Reading them lets a language picked in either app hold in both. Cells' downloads
      (`app/src/proofing/dictionaries.ts` + `dictionaries-data.ts`, made by `scripts/build-dictionaries.mjs`: wooorm's
      dictionary-* packages pinned on jsDelivr with SHA-256, kept in IndexedDB, read by nspell in a worker) can be
      reused as they are; Italian and Portuguese are left out because nspell can't read them.
