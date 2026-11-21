/**
 * What Help › Help (F1) shows: Cells' special features (the same ones the feature hints introduce, with where to find
 * them) and its keyboard shortcuts, all kept here so the pane works offline. Show Training walks through the features
 * in this order.
 */
import { FEATURE_HINTS } from "@/hints/feature-hints"
import { controlIds, TABS } from "@/ribbon/tabs"

export type FeatureTopic = {
  /** The feature hint's id: showing the feature from Help counts as having seen its hint. */
  id: string
  title: string
  text: string
  /** The control "Show me" points at, by its accessible name (as the hints do). */
  target: string
  /** The ribbon command, for the tab to show before pointing. */
  command?: string
  /** Where it is, for features with no ribbon button of their own (shown instead of the tab and group). */
  where?: string
  /** Words people may search with. */
  keywords: string
}

/** The ribbon command behind each hint, and words to find it by. */
const EXTRA: Record<string, { command?: string; where?: string; keywords: string }> = {
  "quick-analysis": {
    where: "The button at the bottom-right corner of a selection of data, or Ctrl+Q",
    keywords: "totals chart sparklines format table preview",
  },
  "flash-fill": { command: "QuickCode", keywords: "split names pattern column fill example" },
  "ai-tools": { command: "AIToolsMenu", keywords: "ai formula explain summarize chart copilot assistant" },
  "recommended-pivots": { command: "PivotTableSuggestion", keywords: "pivot pivottable summary suggest" },
  "describe-pivot": {
    where: "The PivotTable Fields pane, with AI turned on (File › Options › AI)",
    keywords: "ai pivot pivottable words layout",
  },
  "forecast-sheet": { command: "ForecastInsert", keywords: "forecast trend predict future dates" },
  "3d-models": { command: "Insert3DModelDropdown", keywords: "3d model glb turn rotate" },
  "new-window": { command: "WindowNew", keywords: "window arrange side by side compare" },
  ink: { command: "InkPen0", keywords: "draw pen highlighter ink stylus handwriting annotate eraser" },
  "power-query": {
    command: "PowerQueryGetTransformDataMenuSubscription19",
    keywords: "power query get data from web table range query editor import csv json refresh connection transform",
  },
  "chart-format": {
    where: "Format, beside Chart Design (select a chart first)",
    keywords:
      "chart format part element title axis series legend fill outline colour color text font plot area gridlines",
  },
  "header-footer-page": {
    command: "HeaderFooterInsert",
    keywords: "header footer page layout view page number date sheet name file print top bottom margin",
  },
  "chart-sheets": {
    where: "Chart Design › Location (select a chart first)",
    keywords: "chart sheet move chart f11 new sheet object in full window",
  },
}

export const FEATURE_TOPICS: FeatureTopic[] = FEATURE_HINTS.map((hint) => ({
  id: hint.id,
  title: hint.title.replace(/^New: /, ""),
  text: hint.text,
  target: hint.target,
  ...EXTRA[hint.id],
  keywords: EXTRA[hint.id]?.keywords ?? "",
}))

/** "Home › Editing": the tab and group a ribbon command is in. */
export function placeOf(command: string): string | undefined {
  for (const tab of TABS) {
    if (!controlIds(tab).includes(command)) continue
    const group = tab.groups.find((g) => controlIds({ ...tab, groups: [g] }).includes(command))
    return group ? `${tab.label} › ${group.label}` : tab.label
  }
  return undefined
}

/** The tab a command's button is on. */
export function tabOf(command: string): string | undefined {
  return TABS.find((tab) => controlIds(tab).includes(command))?.id
}

export const whereIs = (topic: FeatureTopic) => topic.where ?? (topic.command ? placeOf(topic.command) : undefined)

export type Shortcut = { keys: string; does: string }
export type ShortcutGroup = { title: string; shortcuts: Shortcut[] }

/** Cells' keyboard shortcuts, as Excel has them (each one handled in `sheet/controller.ts`, App.tsx or the ribbon). */
export const SHORTCUTS: ShortcutGroup[] = [
  {
    title: "Moving around",
    shortcuts: [
      { keys: "Arrow keys", does: "Move one cell" },
      { keys: "Ctrl+Arrow", does: "Go to the edge of the data" },
      { keys: "Home", does: "Go to the start of the row" },
      { keys: "Ctrl+Home", does: "Go to A1" },
      { keys: "Ctrl+End", does: "Go to the last used cell" },
      { keys: "PageDown / PageUp", does: "Move one screen down or up" },
      { keys: "Alt+PageDown / Alt+PageUp", does: "Move one screen right or left" },
      { keys: "Ctrl+PageDown / Ctrl+PageUp", does: "Go to the next or previous sheet" },
      { keys: "Ctrl+G or F5", does: "Go To" },
      { keys: "Ctrl+Backspace", does: "Scroll back to the active cell" },
      { keys: "Ctrl+F6", does: "Go to the workbook's next window" },
    ],
  },
  {
    title: "Selecting",
    shortcuts: [
      { keys: "Shift+Arrow", does: "Extend the selection" },
      { keys: "Ctrl+Shift+Arrow", does: "Extend the selection to the edge of the data" },
      { keys: "Ctrl+A", does: "Select the data around the cell, then the whole sheet" },
      { keys: "Ctrl+Shift+*", does: "Select the current region" },
      { keys: "Shift+Space", does: "Select whole rows" },
      { keys: "Ctrl+Space", does: "Select whole columns" },
      { keys: "Ctrl+Shift+PageDown / PageUp", does: "Group the next or previous sheet with this one" },
    ],
  },
  {
    title: "Entering and editing",
    shortcuts: [
      { keys: "F2", does: "Edit the active cell" },
      { keys: "Enter / Shift+Enter", does: "Finish and move down or up" },
      { keys: "Tab / Shift+Tab", does: "Finish and move right or left" },
      { keys: "Ctrl+Enter", does: "Finish and stay in the cell (fills a whole selection)" },
      { keys: "Alt+Enter", does: "A new line inside the cell" },
      { keys: "Esc", does: "Cancel the entry" },
      { keys: "F4", does: "While typing a formula: switch between $A$1, A$1, $A1 and A1" },
      { keys: "Delete", does: "Clear the contents" },
      { keys: "Ctrl+D / Ctrl+R", does: "Fill down or right" },
      { keys: "Ctrl+E", does: "Flash Fill" },
      { keys: "Ctrl+Z / Ctrl+Y", does: "Undo or redo" },
      { keys: "Ctrl+X / Ctrl+C / Ctrl+V", does: "Cut, copy, paste" },
      { keys: "Ctrl++ / Ctrl+-", does: "Insert or delete cells" },
      { keys: "Alt+Down", does: "Open the filter button or the list of a cell" },
    ],
  },
  {
    title: "Formatting",
    shortcuts: [
      { keys: "Ctrl+1", does: "Format Cells" },
      { keys: "Ctrl+B / Ctrl+I / Ctrl+U", does: "Bold, italic, underline" },
      { keys: "Ctrl+5", does: "Strikethrough" },
      { keys: "Ctrl+Shift+$", does: "Currency" },
      { keys: "Ctrl+Shift+%", does: "Percentage" },
      { keys: "Ctrl+Shift+!", does: "Number with two decimals" },
      { keys: "Ctrl+Shift+#", does: "Date" },
      { keys: "Ctrl+Shift+@", does: "Time" },
      { keys: "Ctrl+Shift+~", does: "General" },
      { keys: "Ctrl+Shift+& / Ctrl+Shift+_", does: "Add or remove the outline border" },
      { keys: "Ctrl+9 / Ctrl+0", does: "Hide rows or columns" },
    ],
  },
  {
    title: "Formulas and data",
    shortcuts: [
      { keys: "Shift+F3", does: "Insert Function" },
      { keys: "Ctrl+F3", does: "Name Manager" },
      { keys: "Ctrl+`", does: "Show formulas" },
      { keys: "F9", does: "Calculate now" },
      { keys: "Ctrl+Alt+F9", does: "Calculate everything again" },
      { keys: "Ctrl+T", does: "Make a table" },
      { keys: "Ctrl+Shift+L", does: "Turn filters on or off" },
      { keys: "Ctrl+Q", does: "Quick Analysis" },
      { keys: "Alt+F1", does: "Insert a chart of the data" },
      { keys: "Alt+Shift+Right / Left", does: "Group or ungroup rows or columns" },
      { keys: "Ctrl+Alt+F5", does: "Refresh All" },
    ],
  },
  {
    title: "Workbook and window",
    shortcuts: [
      { keys: "Ctrl+N / Ctrl+O / Ctrl+S", does: "New, open, save" },
      { keys: "F12", does: "Save As" },
      { keys: "Ctrl+W", does: "Close the workbook" },
      { keys: "Ctrl+P", does: "Print" },
      { keys: "Ctrl+F / Ctrl+H", does: "Find, Replace" },
      { keys: "Ctrl+K", does: "Insert a link" },
      { keys: "Shift+F2 / Ctrl+Shift+F2", does: "A note, a comment" },
      { keys: "F7 / Shift+F7", does: "Spelling, Thesaurus" },
      { keys: "Shift+F11", does: "Insert a sheet" },
      { keys: "Alt", does: "Show the ribbon's KeyTips" },
      { keys: "Alt+Q", does: "Search for a command" },
      { keys: "Ctrl+F1", does: "Show or hide the ribbon" },
      { keys: "F1", does: "Help" },
    ],
  },
]

/** Shortcut keys as written on this computer (⌘ for Ctrl on macOS, as Excel for Mac writes them). */
export function keysText(keys: string, mac: boolean): string {
  if (!mac) return keys
  return keys
    .replace(/Ctrl\+/g, "⌘")
    .replace(/Shift\+/g, "⇧")
    .replace(/Alt\+/g, "⌥")
}

export type HelpResults = { topics: FeatureTopic[]; shortcuts: ShortcutGroup[] }

/** The features and shortcuts with every typed word somewhere in them. */
export function searchHelp(query: string): HelpResults {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (!words.length) return { topics: FEATURE_TOPICS, shortcuts: SHORTCUTS }
  const has = (text: string) => {
    const lower = text.toLowerCase()
    return words.every((w) => lower.includes(w))
  }
  return {
    topics: FEATURE_TOPICS.filter((t) => has(`${t.title} ${t.text} ${t.keywords} ${whereIs(t) ?? ""}`)),
    shortcuts: SHORTCUTS.map((g) => ({
      title: g.title,
      shortcuts: g.shortcuts.filter((s) => has(`${s.keys} ${s.does} ${g.title}`)),
    })).filter((g) => g.shortcuts.length > 0),
  }
}
