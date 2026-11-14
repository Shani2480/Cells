# Verify list

Built, but not yet checked against Excel or on a device.

## Against Excel

- [ ] Grid colours (gridlines, header fill and text, selected-header tint and bar, selection tint): taken from
      memory; the VM's screen showed the lock screen. Compare with a screenshot of Excel's own window, light and dark.
- [ ] General format: the thresholds where small and large numbers switch to scientific in a 64 px column, and how
      many digits fit at 11 characters.
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
- [ ] Theme colour tints in the colour palettes (computed, not read).
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
- [ ] Copy names ("Sheet1 (2)") when the name is already 29–31 characters long.
- [ ] A cut pasted on another sheet: Excel moves the cells and references to them follow; we paste and clear the
      source, so formulas that pointed at the source keep pointing there.
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
      Resize Table dialogs' wording. The gallery thumbnails next to Excel's. The name a copied sheet's table gets
      (Cells writes Sales_2). The style colours, structured references and the file round trip (with a filtered
      column) were checked in Excel and the Open XML SDK.

## On a device

- [ ] Scrolling frame times in the Tauri window with GPU compositing (the 16 ms p95 budget; only measured on
      software rendering so far).
- [ ] The .deb and AppImage on a second distribution.
- [ ] Touch: pan and the larger hit targets under `pointer: coarse` (Android tablet, Phase 6).
- [ ] Touch gestures on a real tablet: grippers, long press (500 ms), double tap, pinch, and the editor staying above
      the keyboard. Only driven with synthetic touch events in WebKitGTK so far.

## Not built in Phase 1 (noted for later phases)

- Dragging a selection's border to move it (Phase 3, with drag-and-drop).
- Wrap text and automatic row height (Phase 3, with formatting).
- The ribbon (Phase 3).
