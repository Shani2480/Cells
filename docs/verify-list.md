# Verify list

Built, but not yet checked against Excel or on a device.

## Against Excel

- [ ] Grid colours (gridlines, header fill and text, selected-header tint and bar, selection tint): taken from
      memory; the VM's screen showed the lock screen. Compare with a screenshot of Excel's own window, light and dark.
- [x] General format: the thresholds where small and large numbers switch to scientific in a 64 px column, and how
      many digits fit at 11 characters. (2026-11-26: 17 values in a default-width column match Excel 365's text.)
- [ ] Tab-return when Tab moves without typing (only checked with entries typed).
- [ ] Insert dialog default choice for a tall selection (shift right vs down).
- [ ] Auto-fit column width with Aptos Narrow metrics (we measure with the font the system has: Carlito on Linux).
- [ ] Row header width steps for 5- to 7-digit row numbers.
- [ ] Status bar figures: Average/Sum shown in General at 11 characters ("44.56333333"); check large sums and
      formatted active cells.

- [ ] Reference colours (the eight in order) and the box style of pointed references, light and dark.
- [ ] Function list: which functions Excel lists for a prefix and in what order (we list alphabetically, 12 at most).
- [ ] Show Formulas (Ctrl+`): Excel also widens the columns; we don't yet.
- [ ] Status bar in manual mode: when exactly "Calculate" appears and goes.

## Against Excel (Phase 3)

Ribbon and formatting:

- [ ] Accounting drop-down codes for £, €, ¥ and CHF (built by analogy with the $ code; only $ was read from Excel).
- [x] Theme colour tints in the colour palettes (computed, not read). (2026-11-26: every palette tint of the ten theme
      colours matches Excel 365's; 158 of 160 amounts tried match exactly, the two others are shades the palette
      doesn't show, one step off.)
- [ ] KeyTip letters on every tab (from memory; Home's main ones were checked).
- [ ] Center (and the other alignment buttons) toggling back to General on a second press (ExecuteMso didn't
      toggle in the VM).
- [ ] Ctrl+Shift+$ format (`$#,##0.00_);($#,##0.00)`).
- [ ] Format Cells › Fill › Pattern style and colour: not built yet.

Sheet tabs:

- [ ] Tab look: active tab fill, underline and bold name; inactive coloured tabs filled with their colour and the
      text colour on them (we switch to white below a luminance of 140).
- [ ] Tab context menu "Insert...": Excel opens the Insert dialog (Worksheet, Chart, macro sheets); we insert a
      worksheet at once.
- [ ] Tab Color from the tab menu: Excel opens a submenu; we open the palette at the click.
- [x] Copy names ("Sheet1 (2)") when the name is already 29–31 characters long. (2026-11-26: same as Excel 365, the
      name cut to fit " (2)" in 31 characters.)
- [x] A cut pasted on another sheet: Excel moves the cells and references to them follow. (2026-11-26: fixed and
      checked against Excel 365: references follow with the new sheet's name, a moved formula keeps pointing at its
      old sheet unless it moved with the cells it points at, references to overwritten cells become #REF!.)
- [ ] Clicking another sheet's tab while typing a formula: Excel points at the other sheet; we finish the entry first.
- [ ] Not built: grouped sheets (Shift/Ctrl+click on tabs, Select All Sheets), View Code, Protect Sheet, F11 chart
      sheet, the Activate dialog (we list the sheets in a menu on right-click of the scroll arrows). (Protect Sheet is
      built now.)

Phase 3, added at the end:

- [ ] Review › Notes menu control ids (ReviewNewNote, ReviewShowOrHideNote, …): guessed, not read from Excel.
- [ ] Freeze Panes menu ids and the CF preset colours (light red fill with dark red text etc.): from memory.
- [ ] "Workbook is protected and cannot be changed." (our message for a structure change; Excel mostly greys the
      commands out, which we do too).
- [ ] Protect Workbook's ribbon id (ReviewRestrictEditing, taken from the labels file) and whether the button shows
      pressed while protected.
- [ ] Comment card layout and the note box size (144 × 79 px) against Excel's.

## Against Excel (Phase 4, files)

- [ ] The File views (Home, New, Open, Save As, Info, Recent, Export) next to Excel's backstage, and the wording of
      the save prompt, Document Recovery and the open errors.
- [ ] A 50 MB workbook opened in the Tauri window: the UI stays responsive and the progress shows. Only measured in
      Node so far (1M cells, 6.4 MB: open 2.2 s, save 2.0 s).
- [ ] Open and Save in the browser build without the File System Access API (Firefox, Safari): Open uses a file
      input and every save downloads a copy.

## Against Excel (Phase 4, data tools)

- [ ] The filter dropdown's look next to Excel's (built from memory of Excel 365, no reference screenshot): widths,
      the search box, the value list and the order of the submenus.
- [ ] The filter button's funnel and sort-arrow glyphs at 100% and 200%.
- [ ] Filter by Color with "No Fill" saved from Cells (written as a fill of pattern none; Excel's own form for it
      wasn't captured, Excel reapplies it correctly).
- [ ] The Sort dialog's layout against Excel's.
- [ ] Text to Columns: the wizard's three steps against Excel's Convert Text to Columns Wizard, and its messages
      (one column only, nothing to parse, replace the destination). The splitting was checked against Excel.
- [ ] Remove Duplicates: the wording of the result message and the warning, and the dialog's layout, against
      Excel's (the rules for which values match were checked).
- [ ] Data Validation: the dialog's three tabs, the messages for bad values (blank, not a number, Maximum below
      Minimum) and the "extend to these cells" / "more than one type" questions against Excel's wording (written
      from memory). The in-cell list's look and the input message's place next to Excel's. Which entries each rule
      accepts was checked against Excel; the file round trip was checked in Excel and the Open XML SDK.
- [ ] Tables: the Table Design tab (built from memory of Excel 365, no reference screenshot): its groups, the Table
      Name box, the order of the style options, and Quick Styles against Excel's in-ribbon gallery. The Create Table and
      Resize Table dialogs' wording. The gallery thumbnails next to Excel's. The name a copied sheet's table gets is
      Excel's since 2026-11-26 (the name and the new table's id + 1: Table1 → Table13). The style colours, structured references and the file round trip (with a filtered
      column) were checked in Excel and the Open XML SDK.
- [ ] Flash Fill: ~~its results next to Excel's on the same examples~~ (2026-11-26: all 12 examples give Excel 365's
      results; results are now entered as if typed, numbers as numbers, as Excel does), and the wording of the "no pattern" message (from memory). Excel's live suggestion
      while typing (grey preview, Enter accepts) and the Flash Fill Options button aren't built.
- [ ] Charts: the Chart Design tab, its galleries (Quick Layout, Change Colors, Chart Styles) and the Insert galleries'
      sections were built from memory of Excel 365, not checked against screenshots. When the data has as many rows as
      columns, the series run across the rows (checked in Excel 365 on 3×3 and 4×4 blocks, 2026-11-26). Charts Cells writes
      opened in Excel with no repair and the right =SERIES formulas, and an untouched Excel chart is written back as it
      came; how Excel draws a Cells chart (colours, fonts) next to Cells' picture is unchecked. Not built yet: the Format
      tab, the Insert Chart dialog (All Charts, the "More … Charts" rows), Move Chart, chart sheets, and charts the chart
      package can't read (kept in the file, not shown).
- [ ] Printing: the Page Layout tab's Margins and Size menus, the Page Setup dialog's four tabs and the Custom Header
      dialog, and File › Print's settings were built from memory of Excel 365 (no reference screenshots). The Header
      and Footer preset lists are from memory. Print settings Cells writes were read back by Excel (orientation, paper,
      scale, fit-to, margins, centring, gridlines, headings, black and white, page order, first page number, print area,
      titles, header, footer, breaks) and pass the Open XML SDK; Excel's page counts match on the fixture. Excel prints
      columns slightly narrower than the screen's widths, so its fit-to scale can differ (49% in Excel, 53% in Cells, on
      the fixture's Fit sheet). Excel's own PDF of the fixture came out on Letter (the VM's printer), so page placement
      against Excel on A4 is unchecked. PDF pages are pictures (no selectable text). Not built: Page Layout view, Page
      Break Preview, even/first-page headers (kept in the file), header pictures and fonts, Black and white in the
      drawing, print quality and comments/errors as printed options. The printer and copies are chosen in the system's
      print dialog.

## Against Excel (Phase 5)

- [ ] Regional formats against Excel on Windows with a German and a British region: typed entries (1,5 / 1.234 /
      1.5 / 31.1.2026 / 15/1), number formats, General, the short date (format 14) and the formula bar. Built from
      Excel's documented behaviour, not run in a localized Excel. Known differences: Cells keeps formulas in en-US form
      (Excel uses `;` and `,` in a decimal-comma region), TEXT/VALUE/FIXED/DOLLAR work as in en-US Excel, month names are
      English only, the formula bar shows times with AM/PM, and CSV files are read and written with en-US separators
      (Excel uses the system's list and decimal separators).
- [ ] AutoFilter criteria read from an Excel file are parsed with the person's region (Excel stores them in en-US);
      a filter on "1.5" from a file is a date in a German setting.
- [ ] Pasting a copied range and chart into Inkspire, Word and LibreOffice (only the HTML was checked, in tests):
      the table's widths, borders and fills; the chart as an editable Inkspire chart and as a picture elsewhere.
      Conditional formats and table styles aren't in the copied HTML yet.
- [ ] Linked tables and Edit Data with Inkspire, once it has its side (I5, I7): a linked table follows edits in Cells,
      and a chart's data edited in Cells comes back into the document. Only Cells' files were checked so far.
- [ ] The AI tools with a real model (only a fake local endpoint and canned replies were used): Ollama and LM Studio
      on this computer, and one cloud provider with a key entered in Inkspire; a formula, an explanation, a chart and a
      summary each make sense, and Stop works mid-reply.
- [x] In the real build on Linux (2026-11-16, release build on Xvfb): a second launch with a path or a `file://`
      URL, and a `cells://` link with a range, open in the running Cells. Windows and macOS: see below.
- [ ] The shared packages inside Inkspire (I8): its test suite and `.docx` corpus with `@cells/ooxml` and
      `@cells/chart` in place of its own files, and the Spreadsheet block with `calculateGrid`. So far they were only
      compiled against Inkspire's `tsconfig.json`.
- [ ] The first-run setup and File › Options next to Excel's (the setup has no Excel counterpart; Options was built
      from memory of Excel 365's General and Advanced pages, with only the settings Cells has).

## Against Excel (Phase 7, views and windows)

- [ ] Grouped sheets: which commands Excel greys out while sheets are grouped (Cells' list is in
      `app/src/ribbon/registry.ts`, from memory), and the message wording for a refused one (Cells' own).
- [ ] Grouped sheets: whether Excel's Merge, Clear All and inserted rows on a group match Cells' (each sheet's own
      cells are cleared and merged); the look of grouped tabs in dark mode.
- [ ] Split panes in an Excel file: ~~whether xSplit/ySplit count the row and column headings~~ (they do: 26 px and
      20 px at 100%, read and written as Excel does since 2026-11-26), and the split bar's look and width.
- [ ] Split: overlays (charts, notes, arrows, the cell editor) in the top and left panes are placed in the bottom-right
      pane when it shows the same cells; Excel draws them in each pane.
- [ ] Page Break Preview against Excel: the grey and blue, the "Page N" size, and what Excel does when an automatic
      break is dragged further than a page holds (Cells, like Excel, lowers Adjust to until the page holds it: check
      Excel's exact percentage, and what it does when the sheet is fitted to pages).
- [ ] Spelling: where Excel starts and stops with one cell selected (from the active cell, wrapping round with a
      "continue from the beginning?" question?), and its exact end message.
- [ ] Check Accessibility: Excel's exact list of rules and their names (does it flag unclear link text, tables
      without a header row as errors, hidden sheets?), and which cells it counts for contrast.
- [ ] Allow Edit Ranges: Excel's wording for a bad reference, a missing or taken title; whether ranges unlocked
      with a password lock again when the sheet is protected again.
- [ ] Page Layout: the gap between pages, how many blank pages Excel adds after the data, where the ruler counts
      from (Cells: 0 at the margin), and when "Add header" shows (Excel: on hover). Header & Footer on the page:
      Excel's undo label, whether Esc drops the typed text, and the boxes' size and border.
- [ ] Custom Views: Excel's own messages for replacing and deleting a view (Cells' wording), and whether Excel
      greys Custom Views out while sheets are grouped.
- [x] Fill Across Worksheets on whole rows and columns. (2026-11-26: Excel 365 copies the row's format and height too;
      Cells does now.)
- [ ] Shapes: Excel's standard size for a clicked (not dragged) shape and where it lands; the nudge step with and
      without Ctrl; the default look of new WordArt. New shapes and text boxes look as in Excel 365 since 2026-11-26
      (accent 1 with a 15% shade of it as a 1.5 pt outline; text box white with a #BCBCBC 0.75 pt outline). A Cells-made file opens with no
      repair (checked 2026-11-17: shapes, star rotated and flipped, connector, text box, WordArt, picture with a
      border, SVG icon as a Graphic, group).
- [ ] Icons: Excel's icon size and colour on insert (Cells: one inch, black) and the alt text it fills in.
- [ ] Selection Pane: Excel's order and wording; does it list charts' and groups' children?

## Installers (Phase 6)

- [ ] The release workflow's first run: every job green, the Windows `.exe` and macOS `.dmg` install and start (they
      were never built locally; the macOS code path for files from the Finder, `RunEvent::Opened`, wasn't compiled
      here).
- [ ] Android: a stylus on the grid (select, drag the grippers, double-tap to edit), and opening and saving a
      workbook in Google Drive and OneDrive through the picker (only Downloads was tried).
- [x] The phone layout on a real phone (2026-11-16): 120 fps scrolling and typing, open, edit and
      save, the formula keys over its keyboard, Back.
- [x] On the real phone: the palette closes on a swipe from anywhere on its header, the tab picker too (after the fix).
- [x] On the real phone: a 60 ms flick glides on for about 1.5 s; the File view's sections fit as a row that scrolls
      sideways (Save As partly off the edge). The long-press action sheet, and Back closing it.
- [ ] A mid-range phone, if one turns up.
- [ ] The first signed Android release: Inkspire's two secrets copied into this repository, the `.apk` installs over
      an earlier one, and the `.aab` is accepted by Play Console if it is ever used. Only debug builds were made.
- [ ] The macOS menu bar on a Mac (only built, on Linux, to check its shortcuts parse): each item runs; ⌘S runs
      once (the page takes it before the menu); Undo in a dialog's box undoes its text, not the workbook; Quit asks
      to save; ⌘H hides Cells. Excel for Mac's shortcuts were taken from memory: compare with its Help list, especially
      the Control ones (⌃H, ⌃G, ⌃`) and anything else that differs from Windows.
- [ ] Double-clicking a workbook in Explorer, the Finder and a Linux file manager after installing; `.csv` offered in
      Open With (and whether it becomes the default on Linux and Windows).
- [ ] The Flatpak running: it was built locally but not started, because it reads the real `~/.config/inkspire`;
      check settings shared with Inkspire, linked ranges, the keyring and Edit Data copies in `~/.cache/inkspire`.
- [ ] The AppImage on a second distribution (it only registers `cells://` when it runs; file types come from the
      desktop's Open With).

## On a device

- [ ] Scrolling frame times in the Tauri window with GPU compositing (the 16 ms p95 budget; only measured on
      software rendering so far).
- [ ] The .deb and AppImage on a second distribution.
- [x] Touch: pan and the larger hit targets under `pointer: coarse` (a real tablet, 2026-11-16: taps, swipes at
      60 fps, the picker's rows).
- [x] Touch gestures on the real tablet: dragging a gripper stretches the selection, long press opens the menu and
      Back closes it, double tap edits and the editor stays above the keyboard.
- [x] Pinch zoom on a real tablet: checked by hand (2026-11-16).

## Not built in Phase 1 (noted for later phases)

- Dragging a selection's border to move it (Phase 3, with drag-and-drop).
- Wrap text and automatic row height (Phase 3, with formatting).
- The ribbon (Phase 3).
