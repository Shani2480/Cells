import type { Command } from "./registry"

/** Data › Forecast › What-If Analysis. (Scenario Manager and Data Table come with Part 4's next steps.) */
export function whatIfCommands(): Command[] {
  return [
    {
      id: "WhatIfAnalysisMenu",
      label: "What-If Analysis",
      description: "Try out various values for the formulas in the sheet.",
      run: (env) => env.dialog("goalSeek"),
      menu: ["GoalSeek"],
    },
    {
      id: "GoalSeek",
      label: "Goal Seek...",
      description: "Find the right input for the value you want.",
      run: (env) => env.dialog("goalSeek"),
    },
  ]
}
