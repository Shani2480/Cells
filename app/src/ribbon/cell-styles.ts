import {
  CLEAR_ALL,
  OFFICE_THEME,
  themeColorHex,
  themeRefOf,
  type Border,
  type StylePatch,
  type Theme,
  type ThemeColorRef,
} from "@cells/core"

/**
 * Excel's built-in cell styles (Home › Cell Styles), read from Excel 16 with the Office 2023 theme: what each one sets
 * and the gallery section it sits in. Applying a style resets the parts it covers (font, fill, border, number) to the
 * style's values, as Excel does, and records the style's name. Colours that are the theme's (Office's here) and the
 * headings font follow the workbook's theme, as in Excel.
 */
export type CellStyleDef = {
  name: string
  section: "Good, Bad and Neutral" | "Data and Model" | "Titles and Headings" | "Themed Cell Styles" | "Number Format"
  /** Font: colour (undefined = Automatic), size, bold, italic, font name. */
  font?: { color?: string; size?: number; bold?: boolean; italic?: boolean; name?: string }
  /** Fill colour; null = the style sets "no fill". */
  fill?: string | null
  borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border } | null
  numberFormat?: string
}

const thin = (color: string): Border => ({ style: "thin", color })
const box = (b: Border) => ({ top: b, bottom: b, left: b, right: b })

const ACCENT_FILLS = [
  ["#C0E6F5", "#83CCEB", "#44B3E1", "#156082"],
  ["#FBE2D5", "#F7C7AC", "#F1A983", "#E97132"],
  ["#C1F0C8", "#83E28E", "#47D359", "#196B24"],
  ["#CAEDFB", "#94DCF8", "#61CBF3", "#0F9ED5"],
  ["#F2CEEF", "#E49EDD", "#D86DCD", "#A02B93"],
  ["#DAF2D0", "#B5E6A2", "#8ED973", "#4EA72E"],
]

export const CELL_STYLES: CellStyleDef[] = [
  { name: "Normal", section: "Good, Bad and Neutral" },
  { name: "Bad", section: "Good, Bad and Neutral", font: { color: "#9C0006" }, fill: "#FFC7CE" },
  { name: "Good", section: "Good, Bad and Neutral", font: { color: "#006100" }, fill: "#C6EFCE" },
  { name: "Neutral", section: "Good, Bad and Neutral", font: { color: "#9C5700" }, fill: "#FFEB9C" },
  {
    name: "Calculation",
    section: "Data and Model",
    font: { color: "#FA7D00", bold: true },
    fill: "#F2F2F2",
    borders: box(thin("#7F7F7F")),
  },
  {
    name: "Check Cell",
    section: "Data and Model",
    font: { color: "#FFFFFF", bold: true },
    fill: "#A5A5A5",
    borders: box({ style: "double", color: "#3F3F3F" }),
  },
  { name: "Explanatory Text", section: "Data and Model", font: { color: "#7F7F7F", italic: true } },
  {
    name: "Input",
    section: "Data and Model",
    font: { color: "#3F3F76" },
    fill: "#FFCC99",
    borders: box(thin("#7F7F7F")),
  },
  {
    name: "Linked Cell",
    section: "Data and Model",
    font: { color: "#FA7D00" },
    borders: { bottom: { style: "double", color: "#FF8001" } },
  },
  { name: "Note", section: "Data and Model", fill: "#FFFFCC", borders: box(thin("#B2B2B2")) },
  {
    name: "Output",
    section: "Data and Model",
    font: { color: "#3F3F3F", bold: true },
    fill: "#F2F2F2",
    borders: box(thin("#3F3F3F")),
  },
  { name: "Warning Text", section: "Data and Model", font: { color: "#FF0000" } },
  {
    name: "Heading 1",
    section: "Titles and Headings",
    font: { color: "#0E2841", size: 15, bold: true },
    borders: { bottom: { style: "thick", color: "#156082" } },
  },
  {
    name: "Heading 2",
    section: "Titles and Headings",
    font: { color: "#0E2841", size: 13, bold: true },
    borders: { bottom: { style: "thick", color: "#64BEE6" } },
  },
  {
    name: "Heading 3",
    section: "Titles and Headings",
    font: { color: "#0E2841", bold: true },
    borders: { bottom: { style: "medium", color: "#44B3E1" } },
  },
  { name: "Heading 4", section: "Titles and Headings", font: { color: "#0E2841", bold: true } },
  { name: "Title", section: "Titles and Headings", font: { color: "#0E2841", size: 18, name: "Aptos Display" } },
  {
    name: "Total",
    section: "Titles and Headings",
    font: { bold: true },
    borders: { top: thin("#156082"), bottom: { style: "double", color: "#156082" } },
  },
  ...[0, 1, 2].flatMap((level) =>
    ACCENT_FILLS.map((fills, i): CellStyleDef => ({
      name: `${["20%", "40%", "60%"][level]} - Accent${i + 1}`,
      section: "Themed Cell Styles",
      font: {},
      fill: fills[level],
    })),
  ),
  ...ACCENT_FILLS.map((fills, i): CellStyleDef => ({
    name: `Accent${i + 1}`,
    section: "Themed Cell Styles",
    font: { color: "#FFFFFF" },
    fill: fills[3],
  })),
  { name: "Comma", section: "Number Format", numberFormat: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)' },
  { name: "Comma [0]", section: "Number Format", numberFormat: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)' },
  { name: "Currency", section: "Number Format", numberFormat: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
  { name: "Currency [0]", section: "Number Format", numberFormat: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)' },
  { name: "Percent", section: "Number Format", numberFormat: "0%" },
]

/** A style's colour under `theme`, with its theme reference when it is a theme colour. */
export function styleColor(hex: string, theme: Theme = OFFICE_THEME): { hex: string; ref?: ThemeColorRef } {
  const ref = themeRefOf(hex)
  return ref ? { hex: themeColorHex(ref, theme), ref } : { hex }
}

/** The theme's font for a style's font name (the headings font for Office's Aptos Display). */
export function styleFont(name: string | undefined, theme: Theme = OFFICE_THEME): string | undefined {
  return name === OFFICE_THEME.fonts.major ? theme.fonts.major : name
}

function styleBorder(b: Border | undefined, theme: Theme): Border | null {
  if (!b) return null
  if (!b.color) return b
  const c = styleColor(b.color, theme)
  return c.ref ? { style: b.style, color: c.hex, theme: c.ref } : { style: b.style, color: c.hex }
}

/** The patch that applies a style: the parts it covers set to its values (in `theme`), plus its name. */
export function cellStylePatch(def: CellStyleDef, theme: Theme = OFFICE_THEME): StylePatch {
  if (def.name === "Normal") return CLEAR_ALL
  const patch: StylePatch = { cellStyle: def.name }
  if (def.font) {
    const color = def.font.color ? styleColor(def.font.color, theme) : null
    const major = def.font.name === OFFICE_THEME.fonts.major
    Object.assign(patch, {
      fontName: styleFont(def.font.name, theme) ?? null,
      fontScheme: major ? "major" : null,
      fontSize: def.font.size ?? null,
      bold: def.font.bold ?? null,
      italic: def.font.italic ?? null,
      underline: null,
      strike: null,
      fontColor: color?.hex ?? null,
      fontThemeColor: color?.ref ?? null,
    })
  }
  if (def.fill !== undefined) {
    const fill = def.fill ? styleColor(def.fill, theme) : null
    patch.fill = fill?.hex ?? null
    patch.fillThemeColor = fill?.ref ?? null
  }
  if (def.borders !== undefined) {
    patch.borderTop = styleBorder(def.borders?.top, theme)
    patch.borderBottom = styleBorder(def.borders?.bottom, theme)
    patch.borderLeft = styleBorder(def.borders?.left, theme)
    patch.borderRight = styleBorder(def.borders?.right, theme)
  }
  if (def.numberFormat) patch.numberFormat = def.numberFormat
  return patch
}

export const CELL_STYLE_SECTIONS = [
  "Good, Bad and Neutral",
  "Data and Model",
  "Titles and Headings",
  "Themed Cell Styles",
  "Number Format",
] as const
