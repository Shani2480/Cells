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
  ]
}
