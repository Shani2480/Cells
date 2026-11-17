import type { Command } from "./registry"

/** Data › Forecast › What-If Analysis. */
export function whatIfCommands(): Command[] {
  return [
    {
      id: "WhatIfAnalysisMenu",
      label: "What-If Analysis",
      description: "Try out various values for the formulas in the sheet.",
      menu: ["ScenarioManager", "GoalSeek", "DataTable"],
    },
    {
      id: "ScenarioManager",
      label: "Scenario Manager...",
      description:
        "Create and show named sets of values for the cells your formulas use, and compare them in a summary.",
      run: (env) => env.dialog("scenarioManager"),
    },
    {
      id: "GoalSeek",
      label: "Goal Seek...",
      description: "Find the right input for the value you want.",
      run: (env) => env.dialog("goalSeek"),
    },
    {
      id: "DataTable",
      label: "Data Table...",
      description:
        "See the results of a formula for several values at once: put the values in the first row or column of the selection.",
      run: (env) => env.dialog("dataTable"),
    },
  ]
}
