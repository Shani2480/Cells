/** Colours the canvas draws with, read from the CSS tokens so light and dark mode come from one place. */
export type GridTheme = {
  background: string
  gridline: string
  text: string
  headerBackground: string
  headerLine: string
  headerText: string
  headerSelectedBackground: string
  headerSelectedText: string
  headerSelectedBar: string
  headerFullBackground: string
  headerFullText: string
  selectionFill: string
  selectionBorder: string
  fillPreview: string
  cornerTriangle: string
  spillBorder: string
  /** Where frozen rows and columns end. */
  freezeLine: string
  /** The corner marks of a note (red) and a comment (purple). */
  noteMark: string
  commentMark: string
  /** Reference colours for a formula being typed (--ref-0 …). */
  refColors: string[]
}

const TOKENS: Record<Exclude<keyof GridTheme, "refColors">, string> = {
  background: "--grid-bg",
  gridline: "--grid-line",
  text: "--grid-text",
  headerBackground: "--grid-header-bg",
  headerLine: "--grid-header-line",
  headerText: "--grid-header-text",
  headerSelectedBackground: "--grid-header-selected-bg",
  headerSelectedText: "--grid-header-selected-text",
  headerSelectedBar: "--grid-header-selected-bar",
  headerFullBackground: "--grid-header-full-bg",
  headerFullText: "--grid-header-full-text",
  selectionFill: "--grid-selection-fill",
  selectionBorder: "--grid-selection-border",
  fillPreview: "--grid-fill-preview",
  cornerTriangle: "--grid-corner-triangle",
  spillBorder: "--grid-spill-border",
  freezeLine: "--grid-freeze-line",
  noteMark: "--grid-note-mark",
  commentMark: "--grid-comment-mark",
}

/** Light defaults, used when no stylesheet is present (tests). */
export const LIGHT_THEME: GridTheme = {
  background: "#ffffff",
  gridline: "#d4d4d4",
  text: "#000000",
  headerBackground: "#f5f5f5",
  headerLine: "#d4d4d4",
  headerText: "#424242",
  headerSelectedBackground: "#e1f1e7",
  headerSelectedText: "#0e5c2f",
  headerSelectedBar: "#217346",
  headerFullBackground: "#c9e7d4",
  headerFullText: "#0e5c2f",
  selectionFill: "rgb(0 0 0 / 0.09)",
  selectionBorder: "#217346",
  fillPreview: "#7a7a7a",
  cornerTriangle: "#b8b8b8",
  spillBorder: "#4a7bd0",
  freezeLine: "#9e9e9e",
  noteMark: "#e00000",
  commentMark: "#8a3fd1",
  refColors: ["#2e6fd6", "#d13438", "#8a3ffc", "#0f8a3e", "#b83aa8", "#037f87", "#c25200", "#85593a"],
}

export function readGridTheme(element: Element): GridTheme {
  const style = getComputedStyle(element)
  const theme = { ...LIGHT_THEME, refColors: [...LIGHT_THEME.refColors] }
  for (const key of Object.keys(TOKENS) as (keyof typeof TOKENS)[]) {
    const value = style.getPropertyValue(TOKENS[key]).trim()
    if (value) theme[key] = value
  }
  theme.refColors = theme.refColors.map((fallback, i) => style.getPropertyValue(`--ref-${i}`).trim() || fallback)
  return theme
}

/**
 * The cell font: Excel's 11 pt body font. Aptos Narrow (Excel 2024 / 365) where installed, else Calibri, else
 * Carlito (metric-compatible with Calibri, common on Linux).
 */
export const CELL_FONT_FAMILY =
  '"Aptos Narrow", Calibri, Carlito, "Liberation Sans Narrow", "Liberation Sans", sans-serif'
export const CELL_FONT_SIZE_PX = (11 * 96) / 72
export const CELL_FONT = `${CELL_FONT_SIZE_PX}px ${CELL_FONT_FAMILY}`
