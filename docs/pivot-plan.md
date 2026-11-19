# Part 10 — PivotTables, PivotCharts and Timelines (plan)

Phase 7, release 0.7.0. One commit per step, each with tests, an Excel 365 check of the files it writes, and a
release note. Reference files are made in Excel 365 (the VM, `tools/excel-run.py`) and kept in
`packages/io/test/fixtures/excel` (`pivot-*.xlsx`), with Excel's own pictures of them for comparison.

Decided 2026-11-17: build all seven steps in order, the phone layout of the Fields pane with step 4, and the AI
step after step 7.

## Where Cells is now

PivotTables in Excel files are kept whole: the pivot cache, its records and the PivotTable parts are written back
byte for byte, and the cells Excel last wrote are shown as plain values. Nothing in Cells knows they are a
PivotTable, so they don't refresh, can't be changed, and an edit in their cells breaks them.

## Steps

1. **Model and engine (core)** — `core/pivot/`: the cache (fields, shared items, records), the PivotTable definition
   (row, column, filter and value fields, item order and hidden items, layout, styles), and the engine that turns
   both into the cells Excel would write: Sum, Count, Average, Max, Min, Product, Count Numbers, StdDev(P), Var(P);
   subtotals; grand totals; Compact, Outline and Tabular layouts; Repeat Item Labels; blank rows; empty cells and
   error values shown as set. Checked cell by cell against what Excel wrote in the reference files.
2. **Open and keep (io)** — read `pivotCacheDefinition`, `pivotCacheRecords` and `pivotTable` parts into the model;
   write them anew (records included) once a PivotTable changes, as read while it doesn't. The PivotTable's area
   becomes protected the way Excel's is ("We can't change this part of the PivotTable").
   _Done:_ `io/xlsx/pivot.ts` reads caches into `wb.pivotCaches` and PivotTables into `sheet.pivots`, and writes
   them back as read while unchanged, with only the location patched when rows or columns are inserted before them,
   and anew (cache, records, definition with the engine's row and column items) once changed. Checked: the engine's
   layout of every read PivotTable equals the cells Excel saved; all 12 written anew, a changed one, a moved one and
   one made in Cells open in Excel 365 without repair, and Excel's own Refresh leaves every cell as written. Kept as
   read for now (with all PivotTables of that cache): OLAP and external sources, grouped fields, calculated fields
   and items, Show Values As, label and value filters, Top 10 and sort by value — each comes with its step.
   Typing, clearing, pasting or inserting cells through a PivotTable is refused with Excel's message; whole rows and
   columns before it move it, deleting all of its rows or columns deletes it.
3. **Refresh and use**
   _Refresh done:_ core `pivot/refresh.ts` (read the source range, table or name again; shared items keep their
   order with new ones after; each PivotTable keeps its fields by name and items by value; cells laid out anew with
   the value fields' number formats, old ones cleared, one undo step with the "pivotCaches" op), app
   `sheet/pivots.ts` (Refresh, Refresh All, refresh on open, the cell menu). Checked: after changing the data of both
   reference files, Excel's own Refresh gives the same cells as Cells' for all 12 PivotTables. _Styles done:_ core
   `pivot/built-in-styles.ts` (Excel 365's 84 PivotStyle definitions read through COM, like the table styles) and
   `pivot/styles.ts` (`pivotCellLook`: each element over the ones before in Excel's order; outer items are row
   subheadings, across the row in compact and outline form and down their label column in tabular form; only real
   subtotal lines get the subtotal-row element), drawn by the grid under the cells' own formats. Checked against
   Excel's pictures of `pivot-styles.xlsx` (Light16, Light20, Medium2, Medium9, Dark2, Dark8). The style gallery
   comes with the Design tab in step 4. _Expand / collapse done:_ the engine's label cells carry their field, cache item,
   axis, level and +/− button (outer items with a field inside, where the label first shows, unless `showDrill` is
   off); core `pivot/drill.ts` (`withCollapsed`), app `PivotCommands.toggle` / `drill` / `drillState`,
   `grid/PivotButtons.tsx`, double-click, the cell menu's Expand/Collapse. Labels keep their indent in the cells'
   formats, as Excel saves it. Checked: Excel's Refresh of files Cells collapsed gives the same cells for all 12
   PivotTables. _Filter dropdowns done (items, sorts, report filters):_ core `pivot/filter.ts` (`pivotFilterTarget`: the
   compact Row Labels / Column Labels button covers every field of its area, "Select field" picks one; a field's own
   header in outline and tabular form; each report filter's value cell; `pivotChoices`, `withShownItems`, `withSort`
   keeping the items in the order shown so a report filter's pick still points at its item, `withPagePick`,
   `withPageItems` for Select Multiple Items), app `PivotCommands.showItems` / `pickPage` / `sortField`,
   `grid/PivotFilterMenu.tsx`, buttons in `grid/PivotButtons.tsx` (funnel when filtered, arrow when sorted), Alt+Down.
   Checked: every PivotTable of both reference files with an item hidden, fields sorted and report filters picked
   (one item, several) opens in Excel 365 without repair and Excel's Refresh gives the same cells for all 12. Still to
   come in step 3: Label Filters, Value Filters and Top 10 (Excel's `<filters>`, kept as read today), Show Details.
   — Refresh / Refresh All (Data tab, Analyze tab, the cell menu), the filter buttons on the row
   and column labels (Label Filters, Value Filters, Top 10, item checkboxes, Sort), the report filter dropdowns,
   expand / collapse (+/− buttons, Expand / Collapse Entire Field), double-click a value to see its rows on a new
   sheet (Show Details), and the PivotTable styles drawn from the theme (Excel's 85 built-in styles).
4. **Make and change** — Insert › PivotTable (Create PivotTable: table or range, new or existing sheet), Table
   Design › Summarize with PivotTable, the **PivotTable Fields** pane (field list with search, drag into Filters,
   Columns, Rows and Values, Defer Layout Update), Value Field Settings (Summarize Values By, Number Format, custom
   name), Field Settings (subtotals, layout, print), PivotTable Options, the **PivotTable Analyze** and **Design**
   tabs, Change Data Source, Move / Clear / Select PivotTable. Each change one undo step.
5. **Show Values As, calculated fields, grouping** — % of Grand Total / Column / Row / Parent, Difference From,
   Running Total In, Rank; Calculated Field; Group dates (Years, Quarters, Months, Days) and numbers (by ranges),
   Group Selection for text items. `GETPIVOTDATA`, and Generate GetPivotData when clicking a PivotTable cell in a
   formula.
6. **PivotCharts** — Insert › PivotChart and PivotChart from the Analyze tab: a chart that follows its PivotTable
   (field buttons on the chart, filtering from the chart), written with `c:pivotSource`.
7. **Recommended PivotTables, Quick Analysis, Slicers and Timelines** — Recommended PivotTables (Insert tab), the
   PivotTable choices on Quick Analysis' Tables tab, slicers for PivotTables (today a file with them keeps them as
   read), and **Timelines** (Insert › Timeline, the Timeline tab, x15 timeline parts) for date fields.

## Staying out

- PivotTables from the Data Model / Power Pivot (OLAP, "Add this data to the Data Model", Distinct Count via the
  Data Model, Relationships): they need Microsoft's engine. Files with them keep them as read, as today.
- External data sources (other workbooks' connections, SQL): only ranges and tables in the workbook.
- Calculated Items (rarely used, and they block grouping in Excel): kept as read, not made.

## Phone and tablet

The Fields pane is a side pane on desktop and tablet; on a phone it is a full-screen sheet (Fields, then the four
areas below), with tap-to-add and a long press to move a field between areas instead of dragging. Filter dropdowns
use the phone's bottom sheet, as the AutoFilter's do.

## AI (opt-in, uses the AI tools of Part 6)

"Describe the PivotTable you want" in the Fields pane and in Recommended PivotTables: e.g. "sales by region per
quarter, biggest first" is turned into a PivotTable definition (fields, areas, summary, sort) that is shown as a
preview before it is made. It only proposes a definition; the engine of step 1 does the counting, so the numbers
are never made up by the model.

8. **AI** (after step 7, opt-in): "Describe the PivotTable you want", as above.
