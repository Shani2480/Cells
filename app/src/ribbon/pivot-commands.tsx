import {
  pivotReportLayout,
  pivotSubtotalsShown,
  withBlankLines,
  withGrandTotals,
  withRepeatAllLabels,
  withReportLayout,
  withSubtotalsShown,
  type PivotReportLayout,
  type PivotTable,
} from "@cells/core"
import type { Command, RibbonEnv } from "./registry"
import { PivotStyleGallery, PivotStyleThumb } from "./pivot-styles"

const inPivot = (env: RibbonEnv) => env.c.pivots.active() !== null
const table = (env: RibbonEnv): PivotTable | null => env.c.pivots.active()
const alert = (env: RibbonEnv, problem: string | null) => {
  if (problem) env.c.shell.alert(problem)
}

/** A command the PivotTable tabs show that comes in a later step: disabled, as Cells' unbuilt commands are. */
const later = (id: string, label: string, description: string, keytip: string): Command => ({
  id,
  label,
  description,
  keytip,
  enabled: () => false,
})

/** A Layout menu item that is a choice among several (checked when the table shows it). */
const choice = (
  id: string,
  label: string,
  description: string,
  make: (t: PivotTable) => PivotTable,
  shown: (t: PivotTable) => boolean,
): Command => ({
  id,
  label,
  description,
  enabled: inPivot,
  pressed: (env) => {
    const t = table(env)
    return t ? shown(t) : false
  },
  run: (env) => void env.c.pivots.layout(make, label),
})

const styleOption = (
  id: string,
  label: string,
  key: "rowHeaders" | "colHeaders" | "rowStripes" | "colStripes",
  description: string,
  keytip: string,
): Command => ({
  id,
  label,
  description,
  keytip,
  check: true,
  enabled: inPivot,
  pressed: (env) => table(env)?.style[key] ?? false,
  run: (env) => {
    const t = table(env)
    if (t) env.c.pivots.setStyleOption(key, !t.style[key])
  },
})

const layoutChoice = (id: string, label: string, layout: PivotReportLayout, description: string) =>
  choice(
    id,
    label,
    description,
    (t) => withReportLayout(t, layout),
    (t) => pivotReportLayout(t) === layout,
  )

/** The PivotTable Analyze and Design tabs (while the active cell is in a PivotTable). */
export function pivotCommands(): Command[] {
  return [
    // ── PivotTable Analyze ──
    {
      id: "PivotTableName",
      label: "PivotTable Name",
      description: "Edit the name of this PivotTable.",
      keytip: "N",
      enabled: inPivot,
      field: {
        width: 110,
        value: (env) => table(env)?.name ?? "",
        commit: (env, text) => alert(env, env.c.pivots.rename(text)),
      },
    },
    {
      id: "PivotTableOptionsMenu",
      label: "Options",
      description: "Show the PivotTable Options dialog, to set its layout, totals, display and data options.",
      keytip: "T",
      enabled: inPivot,
      run: (env) => env.dialog("pivotOptions"),
      menu: ["PivotTableOptions", "PivotTableShowReportFilterPages", "PivotTableGetPivotData"],
    },
    {
      id: "PivotTableOptions",
      label: "Options...",
      description: "Show the PivotTable Options dialog.",
      enabled: inPivot,
      run: (env) => env.dialog("pivotOptions"),
    },
    later(
      "PivotTableShowReportFilterPages",
      "Show Report Filter Pages...",
      "Make a sheet for each report filter item.",
      "P",
    ),
    {
      id: "PivotTableGetPivotData",
      label: "Generate GetPivotData",
      description: "Write GETPIVOTDATA when you click a PivotTable value in a formula, instead of a cell reference.",
      keytip: "G",
      pressed: (env) => env.c.pivots.generateGetPivotData,
      run: (env) => {
        env.c.pivots.generateGetPivotData = !env.c.pivots.generateGetPivotData
      },
    },
    {
      id: "PivotActiveField",
      label: "Active Field",
      description: "The field or value field at the active cell. Type a name to rename it.",
      keytip: "F",
      enabled: (env) => env.c.pivots.activeField() !== null,
      field: {
        width: 110,
        value: (env) => env.c.pivots.activeField()?.name ?? "",
        commit: (env, text) => alert(env, env.c.pivots.renameActiveField(text)),
      },
    },
    {
      id: "PivotFieldSettings",
      label: "Field Settings",
      description: "Show the settings of the field or value field at the active cell.",
      keytip: "G",
      enabled: (env) => env.c.pivots.activeField() !== null,
      run: (env) => {
        const t = env.c.pivots.activeField()?.target
        if (t) env.dialog(t.kind === "value" ? `pivotValue:${t.index}` : `pivotField:${t.field}`)
      },
    },
    later("PivotDrillDown", "Drill Down", "Show the next level of a Data Model hierarchy.", "DD"),
    later("PivotDrillUp", "Drill Up", "Show the level above in a Data Model hierarchy.", "DU"),
    {
      id: "PivotExpandField",
      label: "Expand Field",
      description: "Show the inner items of every item of the active field.",
      keytip: "A",
      enabled: (env) => env.c.pivots.drillState()?.expandField ?? false,
      run: (env) => void env.c.pivots.drill("field", false),
    },
    {
      id: "PivotCollapseField",
      label: "Collapse Field",
      description: "Hide the inner items of every item of the active field.",
      keytip: "H",
      enabled: (env) => env.c.pivots.drillState()?.collapseField ?? false,
      run: (env) => void env.c.pivots.drill("field", true),
    },
    {
      id: "PivotGroupSelection",
      label: "Group Selection",
      description: "Gather the selected items into a group; dates and numbers open Grouping to group them by ranges.",
      keytip: "GS",
      enabled: (env) => env.c.pivots.groupTarget() !== null,
      run: (env) => void env.c.pivots.groupSelection(),
    },
    {
      id: "PivotUngroup",
      label: "Ungroup",
      description: "Split the selected groups into their items again, or remove a field's ranges or date parts.",
      keytip: "U",
      enabled: (env) => env.c.pivots.canUngroup(),
      run: (env) => void env.c.pivots.ungroup(),
    },
    {
      id: "PivotGroupField",
      label: "Group Field",
      description: "Group the active date or number field: by years, quarters, months or days, or in ranges.",
      keytip: "GF",
      enabled: (env) => env.c.pivots.groupDialogField() >= 0,
      run: (env) => env.dialog("pivotGroup"),
    },
    later("PivotInsertSlicer", "Insert Slicer", "Filter the PivotTable with buttons (comes with slicers).", "SF"),
    later("PivotInsertTimeline", "Insert Timeline", "Filter the PivotTable by dates (comes with timelines).", "IT"),
    later(
      "PivotFilterConnections",
      "Filter Connections",
      "Choose the slicers and timelines that filter this PivotTable.",
      "FC",
    ),
    {
      id: "PivotRefreshMenu",
      label: "Refresh",
      description: "Get the latest data from the source of this PivotTable.",
      shortcut: "Alt+F5",
      keytip: "R",
      enabled: inPivot,
      run: (env) => void env.c.pivots.refresh(),
      menu: ["PivotRefresh", "PivotRefreshAll"],
    },
    {
      id: "PivotRefresh",
      label: "Refresh",
      description: "Get the latest data from the source of this PivotTable.",
      shortcut: "Alt+F5",
      keytip: "R",
      enabled: inPivot,
      run: (env) => void env.c.pivots.refresh(),
    },
    {
      id: "PivotRefreshAll",
      label: "Refresh All",
      description: "Read every PivotTable source in the workbook again.",
      shortcut: "Ctrl+Alt+F5",
      keytip: "A",
      enabled: inPivot,
      run: (env) => void env.c.pivots.refreshAll(),
    },
    {
      id: "PivotChangeDataSource",
      label: "Change Data Source",
      description: "Change the table or range this PivotTable summarizes.",
      keytip: "I",
      enabled: inPivot,
      run: (env) => env.dialog("pivotSource"),
    },
    {
      id: "PivotClearMenu",
      label: "Clear",
      description: "Remove the fields, formatting and filters of this PivotTable, or only its filters.",
      keytip: "C",
      enabled: inPivot,
      menu: ["PivotClearAll", "PivotClearFilters"],
    },
    {
      id: "PivotClearAll",
      label: "Clear All",
      description: "Remove every field and filter of this PivotTable; it stays, empty.",
      keytip: "A",
      enabled: inPivot,
      run: (env) => void env.c.pivots.clear("all"),
    },
    {
      id: "PivotClearFilters",
      label: "Clear Filters",
      description: "Show every item again: no unticked items, label or value filters, report filters at (All).",
      keytip: "F",
      enabled: inPivot,
      run: (env) => void env.c.pivots.clear("filters"),
    },
    {
      id: "PivotSelectMenu",
      label: "Select",
      description: "Select the labels, the values or the whole PivotTable.",
      keytip: "W",
      enabled: inPivot,
      menu: ["PivotSelectLabelsAndValues", "PivotSelectValues", "PivotSelectLabels", "-", "PivotSelectEntire"],
    },
    {
      id: "PivotSelectLabelsAndValues",
      label: "Labels and Values",
      enabled: inPivot,
      run: (env) => void env.c.pivots.select("labelsAndValues"),
    },
    { id: "PivotSelectValues", label: "Values", enabled: inPivot, run: (env) => void env.c.pivots.select("values") },
    { id: "PivotSelectLabels", label: "Labels", enabled: inPivot, run: (env) => void env.c.pivots.select("labels") },
    {
      id: "PivotSelectEntire",
      label: "Entire PivotTable",
      enabled: inPivot,
      run: (env) => void env.c.pivots.select("entire"),
    },
    {
      id: "PivotMove",
      label: "Move PivotTable",
      description: "Move this PivotTable to another place in the workbook.",
      keytip: "M",
      enabled: inPivot,
      run: (env) => env.dialog("pivotMove"),
    },
    {
      id: "PivotCalculatedFields",
      label: "Fields, Items, & Sets",
      description: "Make fields and items worked out by a formula, and sets.",
      keytip: "J",
      enabled: inPivot,
      menu: [
        "PivotCalculatedField",
        "PivotCalculatedItem",
        "PivotSolveOrder",
        "PivotListFormulas",
        "PivotCreateSet",
        "PivotManageSets",
      ],
    },
    {
      id: "PivotCalculatedField",
      label: "Calculated Field...",
      description: "Add a field worked out by a formula over the other fields, such as =Amount/Units.",
      keytip: "F",
      enabled: inPivot,
      run: (env) => env.dialog("pivotCalcField"),
    },
    later("PivotCalculatedItem", "Calculated Item...", "An item worked out from other items of a field.", "I"),
    later("PivotSolveOrder", "Solve Order...", "The order calculated items are worked out in.", "S"),
    later("PivotListFormulas", "List Formulas", "A new sheet listing the calculated fields and items.", "L"),
    later("PivotCreateSet", "Create Set...", "Sets come with data models, which Cells doesn't open.", "C"),
    later("PivotManageSets", "Manage Sets...", "Sets come with data models, which Cells doesn't open.", "M"),
    {
      id: "PivotFieldList",
      label: "Field List",
      description: "Show or hide the PivotTable Fields pane.",
      keytip: "L",
      enabled: inPivot,
      pressed: (env) => env.c.pivots.fieldsOn(),
      run: (env) => env.c.pivots.toggleFields(),
    },
    {
      id: "PivotButtons",
      label: "+/- Buttons",
      description: "Show or hide the buttons that expand and collapse items.",
      keytip: "B",
      enabled: inPivot,
      pressed: (env) => table(env)?.showDrill !== false,
      run: (env) => void env.c.pivots.setButtons(table(env)?.showDrill === false),
    },
    later("PivotFieldHeaders", "Field Headers", "Show or hide the row and column field headers.", "D"),

    // ── Design ──
    {
      id: "PivotSubtotalsMenu",
      label: "Subtotals",
      description: "Show or hide subtotals, at the bottom or the top of each group.",
      keytip: "ST",
      enabled: inPivot,
      menu: ["PivotSubtotalsNone", "PivotSubtotalsBottom", "PivotSubtotalsTop"],
    },
    choice(
      "PivotSubtotalsNone",
      "Do Not Show Subtotals",
      "No subtotals for any field.",
      (t) => withSubtotalsShown(t, "none"),
      (t) => pivotSubtotalsShown(t) === "none",
    ),
    choice(
      "PivotSubtotalsBottom",
      "Show all Subtotals at Bottom of Group",
      "Each group's subtotal on a line after it.",
      (t) => withSubtotalsShown(t, "bottom"),
      (t) => pivotSubtotalsShown(t) === "bottom",
    ),
    choice(
      "PivotSubtotalsTop",
      "Show all Subtotals at Top of Group",
      "Each group's subtotal on its own heading line.",
      (t) => withSubtotalsShown(t, "top"),
      (t) => pivotSubtotalsShown(t) === "top",
    ),
    {
      id: "PivotGrandTotalsMenu",
      label: "Grand Totals",
      description: "Show or hide the grand totals for rows and columns.",
      keytip: "G",
      enabled: inPivot,
      menu: ["PivotGrandTotalsOff", "PivotGrandTotalsOn", "PivotGrandTotalsRows", "PivotGrandTotalsColumns"],
    },
    choice(
      "PivotGrandTotalsOff",
      "Off for Rows and Columns",
      "No Grand Total row or column.",
      (t) => withGrandTotals(t, false, false),
      (t) => !t.colGrandTotals && !t.rowGrandTotals,
    ),
    choice(
      "PivotGrandTotalsOn",
      "On for Rows and Columns",
      "A Grand Total row and column.",
      (t) => withGrandTotals(t, true, true),
      (t) => t.colGrandTotals && t.rowGrandTotals,
    ),
    choice(
      "PivotGrandTotalsRows",
      "On for Rows Only",
      "A Grand Total column at the right, for each row.",
      (t) => withGrandTotals(t, true, false),
      (t) => t.colGrandTotals && !t.rowGrandTotals,
    ),
    choice(
      "PivotGrandTotalsColumns",
      "On for Columns Only",
      "A Grand Total row at the bottom, for each column.",
      (t) => withGrandTotals(t, false, true),
      (t) => !t.colGrandTotals && t.rowGrandTotals,
    ),
    {
      id: "PivotReportLayoutMenu",
      label: "Report Layout",
      description: "Show the PivotTable in compact, outline or tabular form, and repeat item labels or not.",
      keytip: "P",
      enabled: inPivot,
      menu: [
        "PivotLayoutCompact",
        "PivotLayoutOutline",
        "PivotLayoutTabular",
        "-",
        "PivotRepeatLabels",
        "PivotNoRepeatLabels",
      ],
    },
    layoutChoice("PivotLayoutCompact", "Show in Compact Form", "compact", "Row fields share one column, indented."),
    layoutChoice("PivotLayoutOutline", "Show in Outline Form", "outline", "Each row field in its own column."),
    layoutChoice("PivotLayoutTabular", "Show in Tabular Form", "tabular", "A classic table, each field in a column."),
    choice(
      "PivotRepeatLabels",
      "Repeat All Item Labels",
      "Show each item's label on every line it covers.",
      (t) => withRepeatAllLabels(t, true),
      (t) => t.rows.some((f) => f >= 0) && t.rows.every((f) => f < 0 || t.fields[f].repeatLabels),
    ),
    choice(
      "PivotNoRepeatLabels",
      "Do Not Repeat Item Labels",
      "Show each item's label once.",
      (t) => withRepeatAllLabels(t, false),
      (t) => t.rows.every((f) => f < 0 || !t.fields[f].repeatLabels),
    ),
    {
      id: "PivotBlankRowsMenu",
      label: "Blank Rows",
      description: "Put a blank line after each group, or take them away.",
      keytip: "B",
      enabled: inPivot,
      menu: ["PivotBlankRowsInsert", "PivotBlankRowsRemove"],
    },
    choice(
      "PivotBlankRowsInsert",
      "Insert Blank Line after Each Item",
      "A blank line after each group.",
      (t) => withBlankLines(t, true),
      () => false,
    ),
    choice(
      "PivotBlankRowsRemove",
      "Remove Blank Line after Each Item",
      "No blank lines.",
      (t) => withBlankLines(t, false),
      () => false,
    ),
    styleOption("PivotStyleRowHeaders", "Row Headers", "rowHeaders", "Show the first column in a special style.", "R"),
    styleOption(
      "PivotStyleColumnHeaders",
      "Column Headers",
      "colHeaders",
      "Show the header row in a special style.",
      "H",
    ),
    styleOption("PivotStyleBandedRows", "Banded Rows", "rowStripes", "Shade every other row.", "B"),
    styleOption("PivotStyleBandedColumns", "Banded Columns", "colStripes", "Shade every other column.", "C"),
    {
      id: "PivotStylesGallery",
      label: "PivotTable Styles",
      description: "Pick a visual style for the PivotTable.",
      keytip: "S",
      enabled: inPivot,
      largeIcon: (env) => (
        <PivotStyleThumb style={table(env)?.style.name ?? "PivotStyleLight16"} theme={env.c.wb.theme} />
      ),
      menu: [
        {
          panel: (env, close) => {
            const t = table(env)
            return (
              <PivotStyleGallery
                current={t?.style.name ?? null}
                options={t?.style}
                theme={env.c.wb.theme}
                onPick={(style) => {
                  close()
                  env.c.pivots.setStyle(style)
                }}
              />
            )
          },
        },
        "-",
        "PivotTableNewStyle",
        "PivotStyleClear",
      ],
    },
    {
      id: "PivotStyleClear",
      label: "Clear",
      description: "No style: the PivotTable's cells show plain.",
      keytip: "C",
      enabled: inPivot,
      run: (env) => void env.c.pivots.setStyle(null),
    },
  ]
}
