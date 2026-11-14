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

## On a device

- [ ] Scrolling frame times in the Tauri window with GPU compositing (the 16 ms p95 budget; only measured on
      software rendering so far).
- [ ] The .deb and AppImage on a second distribution.
- [ ] Touch: pan and the larger hit targets under `pointer: coarse` (Android tablet, Phase 6).

## Not built in Phase 1 (noted for later phases)

- Dragging a selection's border to move it (Phase 3, with drag-and-drop).
- Wrap text and automatic row height (Phase 3, with formatting).
- The ribbon (Phase 3).
