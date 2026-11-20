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
   come in step 3: nothing (see Show Details below). _Label Filters, Value Filters and Top 10 done:_ `PivotTable.filters` (Excel's
   `<filters>`, written as Excel writes them, `measureFilter` on the field) and `hideValuesRow` (Excel 365's Show the
   Values row off, read from the x14 extension); the engine applies label filters to the records everywhere (like
   unticked items: label text compared as text, numbers too, wildcards in Equals), and value filters and Top 10
   (count, percent, sum; Bottom too) per item within its parent items and only on their own axis, as Excel keeps the
   other axis' items when a column filter leaves them no values. One filter per field (a new one replaces it), app
   `PivotCommands.setFilter` / `clearField`, the dropdown's Label Filters and Value Filters submenus and
   `shell/PivotFilterDialogs.tsx` (Label Filter, Value Filter, Top 10 Filter). Checked: `pivot-filters.xlsx` (15
   PivotTables made in Excel 365 with every kind of filter) lays out as Excel saved it, its filters are written back
   as Excel wrote them, and Excel's Refresh gives the same cells for all 27 PivotTables of the three reference files
   with filters set in Cells. Kept as read: date filters (they come with date grouping in step 5) and "Allow multiple
   filters per field". Not checked against Excel: Top 10 by count when items tie at the cut. _Show Details done:_ core
   `pivotDetailRecords` (value cells carry their row and column line; the records counted in them), app
   `PivotCommands.showDetails` (double-click, the cell menu): a new sheet before the PivotTable's with the rows as
   TableN in TableStyleMedium2, dates in their format, columns fitted, the table selected, one undo step; Enable show
   details off (`enableDrill="0"`, read and written) gives Excel's message. As checked in Excel 365, except the rows'
   order: Cells keeps the source's, Excel's follows its internal index. **Step 3 done.**
   — Refresh / Refresh All (Data tab, Analyze tab, the cell menu), the filter buttons on the row
   and column labels (Label Filters, Value Filters, Top 10, item checkboxes, Sort), the report filter dropdowns,
   expand / collapse (+/− buttons, Expand / Collapse Entire Field), double-click a value to see its rows on a new
   sheet (Show Details), and the PivotTable styles drawn from the theme (Excel's 85 built-in styles).
4. **Make and change** — Insert › PivotTable (Create PivotTable: table or range, new or existing sheet), Table
   Design › Summarize with PivotTable, the **PivotTable Fields** pane (field list with search, drag into Filters,
   Columns, Rows and Values, Defer Layout Update), Value Field Settings (Summarize Values By, Number Format, custom
   name), Field Settings (subtotals, layout, print), PivotTable Options, the **PivotTable Analyze** and **Design**
   tabs, Change Data Source, Move / Clear / Select PivotTable. Each change one undo step.
   _Insert PivotTable and the Fields pane done:_ core `pivot/make.ts` (`newPivotTable` with Excel 365's defaults,
   `withFieldIn` / `withoutField` / `withFieldTicked` / `withValueMoved`; Σ Values joins Columns at two value fields;
   an empty PivotTable keeps 18 × 3 cells), app `PivotCommands.create` / `change`, `CreatePivotDialog`,
   `PivotFieldsPane` (ticks, search, drag and drop, the chips' Move menus), the empty PivotTable's placeholder, and
   Autofit column widths on update (`useAutoFormatting`, read and written; room for filter buttons). Seven
   PivotTables made in Cells open in Excel 365 without repairs and refresh to the same cells. Not yet: Defer Layout
   Update, the Analyze and Design tabs, the pane's phone layout beyond full screen, autofit of PivotTables on other
   sheets when a refresh changes them.
   _Settings dialogs done:_ core `pivot/settings.ts` (`withValueSettings`, `withFieldSettings`, `withPivotOptions`,
   `pivotSettingsTarget`; Excel's name rules: a value field's name can't be a source field's or another value
   field's, a trailing space can; PivotTables on different sheets may share a name), app Value Field Settings (with
   Format Cells' Number tab), Field Settings, PivotTable Options, from the pane's chip menus, the cell menu and a
   header's double-click. `PivotTable.part` keeps what Cells leaves unread with a renamed PivotTable. Five
   PivotTables using every setting open in Excel 365 without repairs, refresh to the same cells, and show the formats,
   names, captions and flags as set. Not yet: Show Values As (step 5), the Print options, Merge and center labels,
   Classic layout, Alt Text.
   _Analyze and Design tabs done:_ core `withSubtotalsShown`, `withGrandTotals`, `withReportLayout`,
   `withRepeatAllLabels`, `withBlankLines`, `withFiltersCleared`, `withAllCleared`; app `PivotCommands` rename,
   active field, Change Data Source, Clear, Select, Move PivotTable (on its sheet, to another or a new one), style and
   style options; `ribbon/pivot-commands.tsx`, `PIVOT_ANALYZE` / `PIVOT_DESIGN`, the PivotTable style gallery
   (thumbnails drawn by the grid's own style code). Five PivotTables laid out by the Design tab open in Excel 365
   without repairs and refresh to the same cells. Greyed out until their steps: grouping (5), calculated fields (5),
   PivotChart (6), slicers, timelines and Recommended PivotTables (7); not built: Show Report Filter Pages, Generate
   GetPivotData, Field Headers, New PivotTable Style.
   _Touch, Defer Layout Update and Delete done:_ a menu on every field of the list (Add to Report Filter / Row
   Labels / Column Labels / Values, also on right-click), menus placed on the window above or below their row,
   finger-sized rows and buttons on coarse pointers, the pane over the whole phone screen (Field List in the phone's
   PivotTable Analyze sheet opens it); Defer Layout Update with Update (changes wait on the pane and lay out as one
   undo step; Update after a change elsewhere drops them); Delete over a selection that holds whole PivotTables
   deletes them, their cells and unused caches, as one undo step (part of one still gets Excel's message). Files
   with one PivotTable deleted, and with all of them deleted, open in Excel 365 without repairs. **Step 4 done.**
5. **Show Values As, calculated fields, grouping** — % of Grand Total / Column / Row / Parent, Difference From,
   Running Total In, Rank; Calculated Field; Group dates (Years, Quarters, Months, Days) and numbers (by ranges),
   Group Selection for text items. `GETPIVOTDATA`, and Generate GetPivotData when clicking a PivotTable cell in a
   formula.
   _Show Values As done:_ all 14 choices (core `pivot/show-as.ts`, applied to each value cell from the summaries of
   the cells it compares with). Excel's rules for empty cells, read from `fixtures/excel/pivot-showas.xlsx` (22
   PivotTables Excel made): the percentages and Index count an empty cell as 0 and leave it blank when the total
   they divide by is empty; % Of and % Difference From give `#NULL!` for an empty cell and blank for an empty base;
   Difference From counts both as 0; running totals show 0; ranks skip empty cells, ties share a rank (1, 1, 2); a
   line that doesn't reach the base field (grand totals, outer levels) is blank. The file keeps them in
   `showDataAs`, or for Excel 2010's choices in the dataField's x14 extension, with `baseItem` a position in the
   base field's items (1048828/1048829 for previous/next; the model keeps the cache item). Value Field Settings'
   Show Values As tab, the cell menu's Show Values As and Summarize Values By submenus, and the Show Values As
   dialog. Excel reads every setting Cells writes and refreshes all 23 PivotTables of the check file to the same
   cells.
   _GETPIVOTDATA done:_ core `pivot/getpivotdata.ts` (`pivotDataCell`, `getPivotDataFormula`), the function in
   `formula/functions/lookup.ts` (it reads the cell it finds and depends on the whole PivotTable, so a refresh or a
   change of layout recalculates it), `EvalContext.pivotAt`. Excel's rules, from 46 formulas typed into Excel over
   the fixtures (`io/test/getpivotdata.test.ts`): names ignore case, a value field by its own or its source's name;
   the items must pick out one shown line per axis (an outer field may be left out when that is unambiguous); a
   report filter counts only when it shows that item; the reference may be any cell of the table or its filters;
   an empty value reads as blank. Generate GetPivotData (Analyze › Options, on by default, kept on the device):
   pointing at a value cell writes the call; cells it can't name (custom subtotals: Excel itself returns #REF! for
   every form) keep a reference. 306 generated calls, one per value cell of two fixtures, give the cell's own value
   in Excel 365.
   _Calculated fields done:_ core `pivot/calculated.ts`. A calculated field is a cache field with a `formula` and no
   records (the file: `<cacheField formula="…" databaseField="0"/>` after the source's columns; its pivotField has
   `dragToRow/Col/Page="0"` and no subtotal), so every PivotTable on the cache shares it. Excel works it out on the
   sums of the fields it names in every cell, subtotal and grand total, whatever the value field's function (Count of
   Price is the price), and a cell whose sums divide by 0 is #DIV/0! (`fixtures/excel/pivot-calc.xlsx`: a ratio, a
   factor, IF, a quoted 'Unit Cost', Count and Average of one). Fields are replaced by placeholders and the formula
   runs through the normal evaluator without cells; references, other names and arrays get Excel's message. Insert
   Calculated Field (Name with Add/Modify, Formula, Delete, Fields with Insert Field); a new field goes into the
   active PivotTable's values. Refresh keeps calculated fields after the source's columns; Show Details leaves them
   out; they can only be values. A file where Cells modified one and added another opens in Excel 365 without
   repairs, Excel reads the five formulas and refreshes all three PivotTables to the same cells. Not yet: calculated
   items, Solve Order, List Formulas.
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
