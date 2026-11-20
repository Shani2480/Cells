/**
 * Hints that introduce Cells' special features, as Office's teaching callouts do (and as Inkspire's do): one at a
 * time, pointing at the feature's own control while it is on screen, worded as news. Only features people may not find
 * on their own; the everyday ones (formatting, sorting, saving) get none. A tap dismisses a hint for good. On by
 * default; File › Options › General turns them off and shows them again.
 *
 * A new special feature gets its hint here when it is added (feature-hints.test.ts checks each target exists).
 */
export type FeatureHint = {
  id: string
  title: string
  text: string
  /** The control it points at, by its accessible name; the hint waits until one is on screen. */
  target: string
}

export const FEATURE_HINTS: FeatureHint[] = [
  {
    id: "quick-analysis",
    title: "New: Quick Analysis",
    text: "Select some data and use this button (or Ctrl+Q) to add formatting, a chart, totals, a table or sparklines in one go, with a preview as you point.",
    target: "Quick Analysis",
  },
  {
    id: "flash-fill",
    title: "New: Flash Fill",
    text: "Type an example or two of what you want next to your data, such as first names from full names, and Flash Fill (Ctrl+E) does the rest of the column.",
    target: "Flash Fill",
  },
  {
    id: "ai-tools",
    title: "New: AI Tools",
    text: "Write a formula from plain words, have one explained, or let AI suggest a chart or summarize your data. It stays off until you turn it on, and you choose which AI to use, one on this computer if you like.",
    target: "AI Tools",
  },
  {
    id: "recommended-pivots",
    title: "New: Recommended PivotTables",
    text: "Not sure how to start a PivotTable? Cells suggests a few for your data, each with a picture of how it looks.",
    target: "Recommended PivotTables",
  },
  {
    id: "describe-pivot",
    title: "New: Describe the PivotTable you want",
    text: 'Say what you\'d like to see, such as "sales by region per quarter, biggest first", and AI lays out the PivotTable for you. Cells works out the numbers.',
    target: "Describe the PivotTable you want…",
  },
  {
    id: "forecast-sheet",
    title: "New: Forecast Sheet",
    text: "Pick a column of dates and the values beside them, and Cells carries them forward on a new sheet, with a chart and a range of likely values.",
    target: "Forecast Sheet",
  },
  {
    id: "3d-models",
    title: "New: 3D Models",
    text: "Put a 3D model in the sheet and drag the handle in its middle to turn it and show it from any side.",
    target: "3D Models",
  },
  {
    id: "new-window",
    title: "New: New Window",
    text: "Show the workbook twice, each with its own sheet and place, and Arrange All lays them side by side inside Cells.",
    target: "New Window",
  },
]

/** The first hint not yet dismissed whose control is on screen, or none. */
export function nextHint(hints: FeatureHint[], dismissed: readonly string[], isShown: (target: string) => boolean) {
  return hints.find((hint) => !dismissed.includes(hint.id) && isShown(hint.target)) ?? null
}
