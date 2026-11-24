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
  {
    id: "ink",
    title: "New: Draw on the sheet",
    text: "Pick a pen or the highlighter and write, circle or mark up right on the sheet, with a stylus, the mouse or (with Draw with Touch) your finger. Review › Hide Ink hides it.",
    target: "Draw with Touch",
  },
  {
    id: "power-query",
    title: "New: Get and shape data",
    text: "Get Data reads a table, a CSV or JSON file on the web, or a web page's tables, and the Query Editor filters, sorts and trims it before it loads. Refresh brings in the latest, and Excel can refresh it too.",
    target: "Get Data",
  },
  {
    id: "chart-sheets",
    title: "New: A chart on a sheet of its own",
    text: "Move Chart puts the chart on a chart sheet, where it fills the window and prints across the page. F11 makes one straight from the selected data, and Excel opens it as a chart sheet too.",
    target: "Move Chart",
  },
  {
    id: "chart-format",
    title: "New: Format any part of a chart",
    text: "Click a chart, then click its title, an axis, a series or the legend to pick it (or choose it in Chart Elements). Shape Fill, Shape Outline and Text Fill colour just that part; double-click it for all its options. Excel keeps the look.",
    target: "Format Selection",
  },
  {
    id: "header-footer-page",
    title: "New: Type headers and footers on the page",
    text: "In Page Layout view, click above or below a page's cells to type its header or footer right there. The Header & Footer tab adds page numbers, dates and the sheet's name.",
    target: "Header & Footer",
  },
  {
    id: "picture-crop",
    title: "New: Crop a picture, or cut it to a shape",
    text: "Select a picture and use Crop: drag the black handles to cut its sides, or drag the picture to show another part. The arrow under Crop cuts it to a circle or any shape, crops to 1:1, 16:9 and more, and fills or fits the frame.",
    target: "Crop",
  },
  {
    id: "picture-corrections",
    title: "New: Corrections, Color and Transparency for pictures",
    text: "Select a picture and use Corrections to sharpen it or change its brightness and contrast, Color to change its saturation and tone or recolour it (grayscale, sepia, a theme colour), and Transparency to let what's behind it show. Set Transparent Color, under Color, clears one colour you click. Excel opens them all as they are.",
    target: "Corrections",
  },
  {
    id: "picture-artistic-effects",
    title: "New: Artistic Effects for pictures",
    text: "Select a picture and use Artistic Effects to make it look like a pencil sketch, a painting, glass, cement and 18 more. Artistic Effects Options sets how strong the effect is and how much of the picture shows through.",
    target: "Artistic Effects",
  },
  {
    id: "pivot-report-pages",
    title: "New: A sheet for each report filter item",
    text: "In a PivotTable with a report filter, the arrow under Options has Show Report Filter Pages: it makes a sheet for each of the filter's items, each with a copy of the PivotTable showing just that item.",
    target: "Options",
  },
  {
    id: "pivot-sort-by-value",
    title: "New: Sort a PivotTable by its values",
    text: "Right-click a value in a PivotTable and pick Sort to order its items by that value, or use More Sort Options to choose the value field, or to keep them in the order you arrange.",
    target: "Field Settings",
  },
  {
    id: "pivot-calculated-items",
    title: "New: Calculated items",
    text: "Fields, Items, & Sets › Calculated Item adds an item worked out from the field's other items, such as =East-West, in every cell and total. Solve Order decides which one counts where two meet, and List Formulas lists them on a sheet.",
    target: "Fields, Items, & Sets",
  },
  {
    id: "pivot-options-layout",
    title: "New: More PivotTable Options",
    text: "PivotTable Options can now merge and center the labels, set the compact indent, lay the report filters out across, sort days and months in calendar order, allow two filters on a field, set print titles and give the PivotTable Alt Text.",
    target: "Options",
  },
  {
    id: "pivot-new-style",
    title: "New: Your own PivotTable styles",
    text: "Make a PivotTable style of your own: PivotTable Styles › New PivotTable Style. Right-click a style to modify, duplicate or delete it, or to make it the default for new PivotTables.",
    target: "PivotTable Styles",
  },
  {
    id: "table-new-style",
    title: "New: Your own table styles",
    text: "Make a table style of your own: Table Design › Quick Styles › New Table Style. Right-click a style to modify, duplicate or delete it, or to make it the default for new tables.",
    target: "Quick Styles",
  },
]

/** The first hint not yet dismissed whose control is on screen, or none. */
export function nextHint(hints: FeatureHint[], dismissed: readonly string[], isShown: (target: string) => boolean) {
  return hints.find((hint) => !dismissed.includes(hint.id) && isShown(hint.target)) ?? null
}
