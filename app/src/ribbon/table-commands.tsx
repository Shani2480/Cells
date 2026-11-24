import { selectRange, tableAt, tableRows } from "@cells/core"
import { customStyleMenu } from "./style-menu"
import type { TableOption } from "@/sheet/tables"
import type { Command, RibbonEnv } from "./registry"
import { FormatAsTableIcon } from "./icons"
import { TableStyleGallery, TableStyleThumb } from "./table-styles"

const inTable = (env: RibbonEnv) => env.c.tables.active() !== null

/** A Table Style Options check box. */
const option = (id: string, label: string, key: TableOption, description: string, keytip: string): Command => ({
  id,
  label,
  description,
  keytip,
  check: true,
  run: (env) => env.c.tables.setOption(key, !env.c.tables.option(key)),
  pressed: (env) => env.c.tables.option(key),
  enabled: inTable,
})

/** Insert › Table, Home › Format as Table and the Table Design tab. */
export function tableCommands(): Command[] {
  return [
    {
      id: "TableInsertExcel",
      label: "Table",
      description:
        "Create a table to organize and analyze related data. Tables make it easy to sort, filter and format.",
      shortcut: "Ctrl+T",
      keytip: "T",
      run: (env) => env.c.tables.openCreate(),
    },
    {
      id: "PivotTableInsert",
      label: "PivotTable",
      description:
        "Easily arrange and summarize complex data in a PivotTable. You can double-click a value to see which detailed values make up the summarized total.",
      keytip: "V",
      run: (env) => env.dialog("createPivot"),
    },
    {
      id: "PivotTableSuggestion",
      label: "Recommended PivotTables",
      description: "Want us to recommend some PivotTables to summarize your data? Just select a cell in your data.",
      keytip: "SP",
      run: (env) => env.dialog("recommendPivots"),
    },
    {
      id: "FormatAsTableGallery",
      label: "Format as Table",
      description: "Quickly convert a range of cells to a table with its own style.",
      keytip: "T",
      largeIcon: <FormatAsTableIcon />,
      menu: [
        {
          panel: (env, close) => (
            <TableStyleGallery
              current={env.c.tables.active()?.style}
              theme={env.c.wb.theme}
              onPick={(style) => {
                close()
                env.c.tables.openCreate(style ?? undefined)
              }}
              custom={env.c.tableStyles.list()}
              menuFor={(style) => customStyleMenu(env, "table", style, close)}
            />
          ),
        },
        "-",
        "TableStyleNew",
        "PivotTableNewStyle",
      ],
    },

    // ── Table Design ──
    {
      id: "TableName",
      label: "Table Name",
      description: "Edit the name of this table to refer to it in formulas.",
      keytip: "A",
      enabled: inTable,
      field: {
        width: 96,
        value: (env) => env.c.tables.active()?.name ?? "",
        commit: (env, text) => void env.c.tables.rename(text.trim()),
      },
    },
    {
      id: "TableResize",
      label: "Resize Table",
      description: "Resize this table by adding or removing rows and columns.",
      keytip: "E",
      enabled: inTable,
      run: (env) => {
        const t = env.c.tables.active()
        if (t) env.c.shell.resizeTable?.(t.range)
      },
    },
    {
      id: "TableRemoveDuplicates",
      label: "Remove Duplicates",
      description: "Delete duplicate rows from the table.",
      keytip: "M",
      enabled: inTable,
      run: (env) => {
        const t = env.c.tables.active()
        if (!t) return
        const rows = tableRows(t)
        const range = { ...t.range, r0: rows.header ?? rows.data[0], r1: rows.data[1] }
        env.c.setSelection(selectRange(range, env.c.selection.active), null)
        env.c.shell.removeDuplicates?.(range)
      },
    },
    {
      id: "TableConvertToRange",
      label: "Convert to Range",
      description: "Convert this table into a normal range of cells. The data and formatting stay.",
      keytip: "G",
      enabled: inTable,
      run: (env) => env.c.tables.convertToRange(),
    },
    option("TableStyleHeaderRow", "Header Row", "header", "Turn the table's header row on or off.", "H"),
    option("TableStyleTotalRow", "Total Row", "totals", "Add a row at the end of the table with totals.", "T"),
    option(
      "TableStyleBandedRows",
      "Banded Rows",
      "rowStripes",
      "Shade even rows differently from odd rows, to make the table easier to read.",
      "R",
    ),
    option(
      "TableStyleFirstColumn",
      "First Column",
      "firstColumn",
      "Show special formatting for the first column.",
      "C",
    ),
    option("TableStyleLastColumn", "Last Column", "lastColumn", "Show special formatting for the last column.", "L"),
    option(
      "TableStyleBandedColumns",
      "Banded Columns",
      "columnStripes",
      "Shade even columns differently from odd columns.",
      "O",
    ),
    option(
      "TableStyleFilterButton",
      "Filter Button",
      "filter",
      "Turn the filter buttons in the header row on or off.",
      "W",
    ),
    {
      id: "TableStylesGallery",
      label: "Quick Styles",
      description: "Pick a visual style for the table.",
      keytip: "S",
      enabled: inTable,
      largeIcon: (env) => (
        <TableStyleThumb
          style={env.c.tables.active()?.style ?? "TableStyleMedium2"}
          theme={env.c.wb.theme}
          custom={env.c.tableStyles.list()}
        />
      ),
      menu: [
        {
          panel: (env, close) => {
            const t = env.c.tables.active()
            return (
              <TableStyleGallery
                current={t?.style ?? null}
                theme={env.c.wb.theme}
                withNone
                options={t ?? undefined}
                onPick={(style) => {
                  close()
                  env.c.tables.setStyle(style)
                }}
                custom={env.c.tableStyles.list()}
                menuFor={(style) => customStyleMenu(env, "table", style, close, (s) => env.c.tables.setStyle(s))}
              />
            )
          },
        },
        "-",
        "TableStyleNew",
        "TableStyleClear",
      ],
    },
    {
      id: "TableSummarizeWithPivot",
      label: "Summarize with PivotTable",
      description: "Summarize this table's data with a PivotTable.",
      keytip: "P",
      run: (env) => {
        const { row, col } = env.c.selection.active
        const table = tableAt(env.c.sheet, row, col)
        if (table) env.dialog(`createPivot:${table.name}`)
      },
    },
    {
      id: "TableStyleNew",
      label: "New Table Style...",
      description: "Make a table style of your own for this workbook.",
      run: (env) => env.dialog("tableStyle:new"),
    },
    {
      id: "PivotTableNewStyle",
      label: "New PivotTable Style...",
      description: "Make a PivotTable style of your own for this workbook.",
      run: (env) => env.dialog("pivotStyle:new"),
    },
    {
      id: "TableStyleClear",
      label: "Clear",
      enabled: inTable,
      run: (env) => env.c.tables.setStyle(null),
    },
  ]
}
