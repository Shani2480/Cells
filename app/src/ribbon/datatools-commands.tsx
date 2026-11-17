import type { Command } from "./registry"

/** Data › Sort & Filter › Advanced, and the Data Tools that come with Part 4. */
export function dataToolCommands(): Command[] {
  return [
    {
      id: "AdvancedFilterDialog",
      label: "Advanced...",
      description: "Filter with a range of conditions you write on the sheet, in place or as a copy somewhere else.",
      keytip: "Q",
      run: (env) => env.dialog("advancedFilter"),
    },
    {
      id: "Consolidate",
      label: "Consolidate...",
      description:
        "Summarise data from separate ranges, on this or other sheets, into one table here, by position or by its labels.",
      keytip: "N",
      run: (env) => env.dialog("consolidate"),
    },
    {
      id: "ForecastInsert",
      label: "Forecast Sheet",
      description:
        "Create a new sheet that predicts where your data over time is going, with a table of forecasts and a chart.",
      keytip: "FC",
      run: (env) => env.dialog("forecast"),
    },
  ]
}
