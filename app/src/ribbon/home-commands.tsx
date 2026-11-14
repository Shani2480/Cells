import { FONT_SIZES, formatValue, formatWithCode, type BorderPreset } from "@cells/core"
import {
  AArrowDown,
  AArrowUp,
  ArrowDownToLine,
  ArrowUpDown,
  Bold,
  ClipboardPaste,
  Copy,
  Eraser,
  Italic,
  Paintbrush,
  Percent,
  Scissors,
  Search,
  Sigma,
  Strikethrough,
  TextAlignCenter,
  TextAlignEnd,
  TextAlignStart,
  Underline,
  WrapText,
  ListIndentDecrease,
  ListIndentIncrease,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
} from "lucide-react"
import { DEFAULT_FONT, NUMBER_FORMATS } from "@/sheet/format"
import { CELL_STYLES, CELL_STYLE_SECTIONS, cellStylePatch } from "./cell-styles"
import { ColorPanel } from "./colors"
import {
  AccountingIcon,
  BorderIcon,
  CellStylesIcon,
  CellsIcon,
  CommaIcon,
  ConditionalFormattingIcon,
  DecreaseDecimalIcon,
  FillColorIcon,
  FontColorIcon,
  FormatAsTableIcon,
  IncreaseDecimalIcon,
  MergeIcon,
  OrientationIcon,
} from "./icons"
import type { Command, RibbonEnv } from "./registry"

const lucide = { size: 16, strokeWidth: 1.6 }
const large = { size: 32, strokeWidth: 1.25 }

/** The fonts the Font box lists (the ones a spreadsheet usually meets); any installed name can be typed. */
export const FONT_NAMES = [
  "Aptos",
  "Aptos Display",
  "Aptos Narrow",
  "Arial",
  "Calibri",
  "Cambria",
  "Candara",
  "Consolas",
  "Courier New",
  "DejaVu Sans",
  "Georgia",
  "Liberation Sans",
  "Liberation Serif",
  "Noto Sans",
  "Segoe UI",
  "Tahoma",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
]

const f = (env: RibbonEnv) => env.c.format
const active = (env: RibbonEnv) => env.c.format.active
const ready = (env: RibbonEnv) => !env.c.edit

/** Number Format box: the category of the active cell's format, as Excel names it. */
export function formatCategory(code: string | undefined): string {
  if (!code || code === "General") return "General"
  if (code === "@") return "Text"
  if (code.includes("* ") && /[$€£¥]|\[\$/.test(code)) return "Accounting"
  if (code.startsWith("_(* ") || code.startsWith("_-* ")) return "Accounting"
  const bare = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, "")
  if (/E[+-]/i.test(bare)) return "Scientific"
  if (bare.includes("%")) return "Percentage"
  if (/\?\/[?\d]/.test(bare)) return "Fraction"
  if (/[dy]/i.test(bare) || /m{3,}/i.test(bare)) return "Date"
  if (/[hs]/i.test(bare) || /AM\/PM/i.test(bare)) return "Time"
  if (/[$€£¥]/.test(code)) return "Currency"
  if (/[0#]/.test(bare)) return "Number"
  return "Custom"
}

const FORMAT_CHOICES: { label: string; code: string; style?: string }[] = [
  { label: "General", code: NUMBER_FORMATS.general },
  { label: "Number", code: "0.00" },
  { label: "Currency", code: '"$"#,##0.00' },
  { label: "Accounting", code: NUMBER_FORMATS.accounting },
  { label: "Short Date", code: NUMBER_FORMATS.shortDate },
  { label: "Long Date", code: NUMBER_FORMATS.longDate },
  { label: "Time", code: NUMBER_FORMATS.longTime },
  { label: "Percentage", code: "0.00%" },
  { label: "Fraction", code: NUMBER_FORMATS.fraction },
  { label: "Scientific", code: NUMBER_FORMATS.scientific },
  { label: "Text", code: NUMBER_FORMATS.text },
]

/** What the active cell's value looks like in a format (the Number Format list previews each). */
function preview(env: RibbonEnv, code: string): string {
  const { row, col } = env.c.selection.active
  const value = env.c.sheet.cells.value(row, col)
  if (value === undefined) return ""
  if (code === "General") return formatValue(value) ?? ""
  if (typeof value === "string" && !code.includes("@")) return value
  if (typeof value === "object") return formatValue(value) ?? ""
  return formatWithCode(value, code).text.replace(/[-]/g, " ").trim()
}

const BORDER_ITEMS: { id: string; preset: BorderPreset; label: string }[] = [
  { id: "BorderBottomNoToggle", preset: "bottom", label: "Bottom Border" },
  { id: "BorderTopNoToggle", preset: "top", label: "Top Border" },
  { id: "BorderLeftNoToggle", preset: "left", label: "Left Border" },
  { id: "BorderRightNoToggle", preset: "right", label: "Right Border" },
  { id: "BorderNone", preset: "none", label: "No Border" },
  { id: "BordersAll", preset: "all", label: "All Borders" },
  { id: "BorderOutside", preset: "outside", label: "Outside Borders" },
  { id: "BorderThickOutside", preset: "thickOutside", label: "Thick Outside Borders" },
  { id: "BorderDoubleBottom", preset: "bottomDouble", label: "Bottom Double Border" },
  { id: "BorderThickBottom", preset: "thickBottom", label: "Thick Bottom Border" },
  { id: "BorderTopAndBottom", preset: "topBottom", label: "Top and Bottom Border" },
  { id: "BorderTopAndThickBottom", preset: "topThickBottom", label: "Top and Thick Bottom Border" },
  { id: "BorderTopAndDoubleBottom", preset: "topDoubleBottom", label: "Top and Double Bottom Border" },
]

/** The colour the Fill / Font Color buttons apply, remembered from the last pick (Excel starts with yellow / red). */
const lastColor = { fill: "#FFFF00", font: "#FF0000" }

/** Whole sheet selections act on the used area for AutoSum's guess. */
function autoSumText(env: RibbonEnv, fn: string): string {
  const { row, col } = env.c.selection.active
  const cells = env.c.sheet.cells
  const isNum = (r: number, c: number) => typeof cells.value(r, c) === "number"
  let r0 = row
  while (r0 > 0 && isNum(r0 - 1, col)) r0--
  if (r0 < row) return `=${fn}(${ref(r0, col)}:${ref(row - 1, col)})`
  let c0 = col
  while (c0 > 0 && isNum(row, c0 - 1)) c0--
  if (c0 < col) return `=${fn}(${ref(row, c0)}:${ref(row, col - 1)})`
  return `=${fn}()`
}

function ref(row: number, col: number) {
  let name = ""
  for (let c = col + 1; c > 0; c = Math.floor((c - 1) / 26)) name = String.fromCharCode(65 + ((c - 1) % 26)) + name
  return `${name}${row + 1}`
}

function autoSum(env: RibbonEnv, fn: string) {
  const text = autoSumText(env, fn)
  env.c.beginEdit("enter", text)
  // Leave the caret inside the brackets when Excel found no range.
  if (text.endsWith("()")) env.c.setEditText(text, text.length - 1)
}

const ALL_TYPES = { numbers: true, text: true, logicals: true, errors: true }

export function homeCommands(): Command[] {
  const cmds: Command[] = [
    // ── Clipboard ──
    {
      id: "PasteMenu",
      label: "Paste",
      description: "Add content from the Clipboard to your sheet.",
      shortcut: "Ctrl+V",
      keytip: "V",
      icon: <ClipboardPaste {...lucide} />,
      largeIcon: <ClipboardPaste {...large} />,
      run: (env) => env.paste(),
      menu: ["Paste", "-", "PasteSpecialDialog"],
    },
    {
      id: "Paste",
      label: "Paste",
      shortcut: "Ctrl+V",
      icon: <ClipboardPaste {...lucide} />,
      run: (env) => env.paste(),
    },
    { id: "PasteSpecialDialog", label: "Paste Special...", shortcut: "Ctrl+Alt+V" },
    {
      id: "Cut",
      label: "Cut",
      description: "Remove the selection and put it on the Clipboard so you can paste it somewhere else.",
      shortcut: "Ctrl+X",
      keytip: "X",
      icon: <Scissors {...lucide} />,
      run: (env) => env.copy(true),
    },
    {
      id: "CopySplitButton",
      label: "Copy",
      description: "Put a copy of the selection on the Clipboard so you can paste it somewhere else.",
      shortcut: "Ctrl+C",
      keytip: "C",
      icon: <Copy {...lucide} />,
      run: (env) => env.copy(false),
      menu: ["Copy", "CopyAsPicture"],
    },
    { id: "Copy", label: "Copy", shortcut: "Ctrl+C", icon: <Copy {...lucide} />, run: (env) => env.copy(false) },
    { id: "CopyAsPicture", label: "Copy as Picture..." },
    {
      id: "FormatPainter",
      label: "Format Painter",
      description:
        "Like the look of a selection? Apply that look to other cells. Pick up the formatting, then select where it goes. Double-click to apply it in several places.",
      keytip: "FP",
      icon: <Paintbrush {...lucide} />,
      run: (env) => f(env).armPainter(),
      pressed: (env) => !!f(env).painter,
    },
    { id: "ShowClipboard", label: "Clipboard", description: "Show the Office Clipboard." },

    // ── Font ──
    {
      id: "Font",
      label: "Font",
      description: "Pick a new font for your text.",
      shortcut: "Ctrl+Shift+F",
      keytip: "FF",
      field: {
        width: 128,
        value: (env) => active(env).fontName ?? DEFAULT_FONT,
        commit: (env, text) => text.trim() && f(env).fontName(text.trim()),
        options: () =>
          FONT_NAMES.map((name) => ({
            label: name,
            value: name,
            preview: <span style={{ fontFamily: `"${name}"`, fontSize: 14 }}>{name}</span>,
          })),
      },
    },
    {
      id: "FontSize",
      label: "Font Size",
      description: "Change the size of your text.",
      shortcut: "Ctrl+Shift+P",
      keytip: "FS",
      field: {
        width: 44,
        value: (env) => String(active(env).fontSize ?? 11),
        commit: (env, text) => {
          f(env).fontSize(Number(text))
        },
        options: () => FONT_SIZES.map((s) => ({ label: String(s), value: String(s) })),
      },
    },
    {
      id: "FontSizeIncrease",
      label: "Increase Font Size",
      description: "Make your text a bit bigger.",
      keytip: "FG",
      icon: <AArrowUp {...lucide} />,
      run: (env) => f(env).stepFont(1),
    },
    {
      id: "FontSizeDecrease",
      label: "Decrease Font Size",
      description: "Make your text a bit smaller.",
      keytip: "FK",
      icon: <AArrowDown {...lucide} />,
      run: (env) => f(env).stepFont(-1),
    },
    {
      id: "Bold",
      label: "Bold",
      description: "Make your text bold.",
      shortcut: "Ctrl+B",
      keytip: "1",
      icon: <Bold {...lucide} strokeWidth={2.2} />,
      run: (env) => f(env).toggle("bold"),
      pressed: (env) => !!active(env).bold,
    },
    {
      id: "Italic",
      label: "Italic",
      description: "Italicize your text.",
      shortcut: "Ctrl+I",
      keytip: "2",
      icon: <Italic {...lucide} />,
      run: (env) => f(env).toggle("italic"),
      pressed: (env) => !!active(env).italic,
    },
    {
      id: "UnderlineGallery",
      label: "Underline",
      description: "Underline your text.",
      shortcut: "Ctrl+U",
      keytip: "3",
      icon: <Underline {...lucide} />,
      run: (env) => f(env).underline(active(env).underline === "double" ? "double" : "single"),
      pressed: (env) => !!active(env).underline,
      menu: ["Underline", "UnderlineDouble"],
    },
    {
      id: "Underline",
      label: "Underline",
      icon: <Underline {...lucide} />,
      run: (env) => f(env).underline("single"),
      pressed: (env) => active(env).underline === "single",
    },
    {
      id: "UnderlineDouble",
      label: "Double Underline",
      icon: <Underline {...lucide} />,
      run: (env) => f(env).underline("double"),
      pressed: (env) => active(env).underline === "double",
    },
    {
      id: "Strikethrough",
      label: "Strikethrough",
      description: "Cross something out by drawing a line through it.",
      shortcut: "Ctrl+5",
      icon: <Strikethrough {...lucide} />,
      run: (env) => f(env).toggle("strike"),
      pressed: (env) => !!active(env).strike,
    },
    {
      id: "BordersGallery",
      label: "Borders",
      description: "Apply borders to the selected cells.",
      keytip: "B",
      icon: (env) => <BorderIcon preset={f(env).lastBorder} />,
      run: (env) => f(env).border(f(env).lastBorder),
      menu: [
        { heading: "Borders" },
        ...BORDER_ITEMS.slice(0, 4).map((b) => b.id),
        "-",
        ...BORDER_ITEMS.slice(4, 8).map((b) => b.id),
        "-",
        ...BORDER_ITEMS.slice(8).map((b) => b.id),
        { heading: "Draw Borders" },
        "BorderDrawLine",
        "BorderDrawGrid",
        "BorderErase",
        "BorderColorPickerExcel",
        "BorderStyle",
        "-",
        "BordersMoreDialog",
      ],
    },
    ...BORDER_ITEMS.map((b): Command => ({
      id: b.id,
      label: b.label,
      icon: <BorderIcon preset={b.preset} />,
      run: (env) => f(env).border(b.preset),
    })),
    { id: "BorderDrawLine", label: "Draw Border" },
    { id: "BorderDrawGrid", label: "Draw Border Grid" },
    { id: "BorderErase", label: "Erase Border" },
    {
      id: "BorderColorPickerExcel",
      label: "Line Color",
      menu: [
        {
          panel: (env, close) => (
            <ColorPanel
              noneLabel="Automatic"
              current={f(env).pen.color}
              onPick={(hex) => {
                f(env).pen = { ...f(env).pen, color: hex ?? undefined }
                close()
              }}
            />
          ),
        },
      ],
    },
    {
      id: "BorderStyle",
      label: "Line Style",
      menu: [
        {
          panel: (env, close) => (
            <div className="line-styles" role="menu">
              {(
                [
                  "hair",
                  "dotted",
                  "dashDotDot",
                  "dashDot",
                  "dashed",
                  "thin",
                  "mediumDashDotDot",
                  "slantDashDot",
                  "mediumDashDot",
                  "mediumDashed",
                  "medium",
                  "thick",
                  "double",
                ] as const
              ).map((style) => (
                <button
                  key={style}
                  type="button"
                  role="menuitemradio"
                  aria-checked={f(env).pen.style === style}
                  aria-label={style}
                  className="line-style-row"
                  onClick={() => {
                    f(env).pen = { ...f(env).pen, style }
                    close()
                  }}
                >
                  <span className={`line-sample is-${style}`} />
                </button>
              ))}
            </div>
          ),
        },
      ],
    },
    {
      id: "BordersMoreDialog",
      label: "More Borders...",
      run: (env) => env.dialog("formatCells:border"),
    },
    {
      id: "CellFillColorPicker",
      label: "Fill Color",
      description: "Color the background of cells to make them stand out.",
      keytip: "H",
      icon: () => <FillColorIcon color={lastColor.fill} />,
      run: (env) => f(env).fill(lastColor.fill),
      menu: [
        {
          panel: (env, close) => (
            <ColorPanel
              noneLabel="No Fill"
              current={active(env).fill}
              onPick={(hex) => {
                if (hex) lastColor.fill = hex
                f(env).fill(hex)
                close()
              }}
            />
          ),
        },
        "FontShadingColorMoreColorsDialog",
      ],
    },
    { id: "FontShadingColorMoreColorsDialog", label: "More Colors..." },
    {
      id: "FontColorPicker",
      label: "Font Color",
      description: "Change the color of your text.",
      keytip: "FC",
      icon: () => <FontColorIcon color={lastColor.font} />,
      run: (env) => f(env).fontColor(lastColor.font),
      menu: [
        {
          panel: (env, close) => (
            <ColorPanel
              noneLabel="Automatic"
              current={active(env).fontColor}
              onPick={(hex) => {
                if (hex) lastColor.font = hex
                f(env).fontColor(hex)
                close()
              }}
            />
          ),
        },
        "FontColorMoreColorsDialogExcel",
      ],
    },
    { id: "FontColorMoreColorsDialogExcel", label: "More Colors..." },
    {
      id: "FormatCellsFontDialog",
      label: "Font Settings",
      description: "Customize your text using advanced font options.",
      shortcut: "Ctrl+Shift+F",
      run: (env) => env.dialog("formatCells:font"),
    },

    // ── Alignment ──
    ...(
      [
        ["AlignTopExcel", "Top Align", "Align text to the top.", "top", "AT", AlignVerticalJustifyStart],
        [
          "AlignMiddleExcel",
          "Middle Align",
          "Align text so that it is centered between the top and bottom of the cell.",
          "center",
          "AM",
          AlignVerticalJustifyCenter,
        ],
        ["AlignBottomExcel", "Bottom Align", "Align text to the bottom.", "bottom", "AB", AlignVerticalJustifyEnd],
      ] as const
    ).map(([id, label, description, align, keytip, Icon]): Command => ({
      id,
      label,
      description,
      keytip,
      icon: <Icon {...lucide} />,
      run: (env) => f(env).vAlign(align),
      pressed: (env) => (active(env).vAlign ?? "bottom") === align,
    })),
    ...(
      [
        ["AlignLeft", "Align Left", "Align your content to the left.", "left", "AL", TextAlignStart],
        ["AlignCenter", "Center", "Center your content.", "center", "AC", TextAlignCenter],
        ["AlignRight", "Align Right", "Align your content to the right.", "right", "AR", TextAlignEnd],
      ] as const
    ).map(([id, label, description, align, keytip, Icon]): Command => ({
      id,
      label,
      description,
      keytip,
      icon: <Icon {...lucide} />,
      run: (env) => f(env).hAlign(align),
      pressed: (env) => active(env).hAlign === align,
    })),
    {
      id: "OrientationMenu",
      label: "Orientation",
      description: "Rotate your text diagonally or vertically. This is a great way to label narrow columns.",
      keytip: "FQ",
      icon: <OrientationIcon />,
      menu: [
        "TextOrientationAngleCounterclockwise",
        "TextOrientationAngleClockwise",
        "TextOrientationVertical",
        "TextOrientationRotateUp",
        "TextOrientationRotateDown",
        "-",
        "CellAlignmentOptions",
      ],
    },
    ...(
      [
        ["TextOrientationAngleCounterclockwise", "Angle Counterclockwise", 45],
        ["TextOrientationAngleClockwise", "Angle Clockwise", 135],
        ["TextOrientationVertical", "Vertical Text", 255],
        ["TextOrientationRotateUp", "Rotate Text Up", 90],
        ["TextOrientationRotateDown", "Rotate Text Down", 180],
      ] as const
    ).map(([id, label, value]): Command => ({
      id,
      label,
      run: (env) => f(env).rotation(value),
      pressed: (env) => active(env).rotation === value,
    })),
    {
      id: "CellAlignmentOptions",
      label: "Format Cell Alignment",
      description: "Customize cell alignment options.",
      run: (env) => env.dialog("formatCells:alignment"),
    },
    {
      id: "IndentDecreaseExcel",
      label: "Decrease Indent",
      description: "Move your content closer to the cell border.",
      shortcut: "Ctrl+Alt+Shift+Tab",
      keytip: "5",
      icon: <ListIndentDecrease {...lucide} />,
      run: (env) => f(env).indent(-1),
    },
    {
      id: "IndentIncreaseExcel",
      label: "Increase Indent",
      description: "Move your content farther away from the cell border.",
      shortcut: "Ctrl+Alt+Tab",
      keytip: "6",
      icon: <ListIndentIncrease {...lucide} />,
      run: (env) => f(env).indent(1),
    },
    {
      id: "WrapText",
      label: "Wrap Text",
      description: "Wrap extra-long text into multiple lines so you can see all of it.",
      keytip: "W",
      icon: <WrapText {...lucide} />,
      run: (env) => f(env).wrap(),
      pressed: (env) => !!active(env).wrap,
    },
    {
      id: "MergeCenterMenu",
      label: "Merge & Center",
      description: "Combine and center the contents of the selected cells in a new larger cell.",
      keytip: "M",
      icon: <MergeIcon />,
      run: (env) => f(env).merge("center"),
      pressed: (env) => f(env).isMerged(),
      menu: ["MergeCenter", "MergeCellsAcross", "MergeCells", "UnmergeCells"],
    },
    { id: "MergeCenter", label: "Merge & Center", icon: <MergeIcon />, run: (env) => f(env).merge("center") },
    { id: "MergeCellsAcross", label: "Merge Across", icon: <MergeIcon />, run: (env) => f(env).merge("across") },
    { id: "MergeCells", label: "Merge Cells", icon: <MergeIcon />, run: (env) => f(env).merge("cells") },
    {
      id: "UnmergeCells",
      label: "Unmerge Cells",
      icon: <MergeIcon />,
      run: (env) => f(env).unmerge(),
      enabled: (env) => f(env).isMerged(),
    },

    // ── Number ──
    {
      id: "NumberFormatGallery",
      label: "Number Format",
      description: "Choose the format for your cells, such as percentage, currency, date or time.",
      keytip: "N",
      field: {
        width: 104,
        value: (env) => formatCategory(active(env).numberFormat),
        commit: (env, text) => {
          const choice = FORMAT_CHOICES.find(
            (c) => c.label.toLowerCase() === text.trim().toLowerCase() || c.code === text,
          )
          if (choice) f(env).numberFormat(choice.code)
          else if (text.trim()) f(env).numberFormat(text.trim())
        },
        options: (env) =>
          FORMAT_CHOICES.map((c) => ({
            label: c.label,
            value: c.code,
            preview: (
              <span className="format-choice">
                <span>{c.label}</span>
                <span className="format-choice-sample">{preview(env, c.code)}</span>
              </span>
            ),
          })),
      },
    },
    {
      id: "NumberFormatsDialog",
      label: "More Number Formats...",
      run: (env) => env.dialog("formatCells:number"),
    },
    {
      id: "AccountingFormatMenu",
      label: "Accounting Number Format",
      description: "Format as dollars, euros or another currency.",
      keytip: "AN",
      icon: <AccountingIcon />,
      run: (env) => f(env).numberFormat(NUMBER_FORMATS.accounting, "Currency"),
      menu: [
        "AccountingFormatEnUs",
        "AccountingFormatEnGb",
        "AccountingFormatEuro",
        "AccountingFormatChina",
        "AccountingFormatSwiss",
        "-",
        "AccountingFormatMoreExcel",
      ],
    },
    ...(
      [
        ["AccountingFormatEnUs", "$ English (United States)", NUMBER_FORMATS.accounting],
        [
          "AccountingFormatEnGb",
          "£ English (United Kingdom)",
          '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
        ],
        [
          "AccountingFormatEuro",
          "€ Euro (€ 123)",
          '_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * "-"??_-;_-@_-',
        ],
        [
          "AccountingFormatChina",
          "¥ Chinese (Simplified, China)",
          '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ',
        ],
        [
          "AccountingFormatSwiss",
          "fr. French (Switzerland)",
          '_-* #,##0.00 [$CHF-100C]_-;-* #,##0.00 [$CHF-100C]_-;_-* "-"?? [$CHF-100C]_-;_-@_-',
        ],
      ] as const
    ).map(([id, label, code]): Command => ({
      id,
      label,
      run: (env) => f(env).numberFormat(code),
    })),
    {
      id: "AccountingFormatMoreExcel",
      label: "More Accounting Formats...",
      run: (env) => env.dialog("formatCells:number"),
    },
    {
      id: "PercentStyle",
      label: "Percent Style",
      description: "Format as a percent.",
      shortcut: "Ctrl+Shift+%",
      keytip: "P",
      icon: <Percent {...lucide} />,
      run: (env) => f(env).numberFormat(NUMBER_FORMATS.percent, "Percent"),
    },
    {
      id: "CommaStyle",
      label: "Comma Style",
      description: "Format with a thousands separator.",
      keytip: "K",
      icon: <CommaIcon />,
      run: (env) => f(env).numberFormat(NUMBER_FORMATS.comma, "Comma"),
    },
    {
      id: "DecimalsIncrease",
      label: "Increase Decimal",
      description: "Show more decimal places for a more precise value.",
      keytip: "0",
      icon: <IncreaseDecimalIcon />,
      run: (env) => f(env).decimals(1),
    },
    {
      id: "DecimalsDecrease",
      label: "Decrease Decimal",
      description: "Show fewer decimal places.",
      keytip: "9",
      icon: <DecreaseDecimalIcon />,
      run: (env) => f(env).decimals(-1),
    },
    {
      id: "FormatCellsNumberDialog",
      label: "Number Format",
      description: "Customize your cell format with advanced options.",
      shortcut: "Ctrl+1",
      run: (env) => env.dialog("formatCells:number"),
    },

    // ── Styles ──
    {
      id: "ConditionalFormattingMenu",
      label: "Conditional Formatting",
      description:
        "Easily spot trends and patterns in your data using bars, colors and icons to visually highlight important values.",
      keytip: "L",
      largeIcon: <ConditionalFormattingIcon />,
      menu: [
        "ConditionalFormattingHighlightCellsMenu",
        "ConditionalFormattingTopBottomMenu",
        "-",
        "ConditionalFormattingDataBarsGallery",
        "ConditionalFormattingColorScalesGallery",
        "ConditionalFormattingIconSetsGallery",
        "-",
        "ConditionalFormattingNewRule",
        "ConditionalFormattingClearMenu",
        "ConditionalFormattingsManage",
      ],
    },
    // The menu's commands live in cf-commands.tsx.
    {
      id: "FormatAsTableGallery",
      label: "Format as Table",
      description: "Quickly convert a range of cells to a table with its own style.",
      keytip: "T",
      largeIcon: <FormatAsTableIcon />,
    },
    {
      id: "CellStylesGallery",
      label: "Cell Styles",
      description: "A colorful style can make important data stand out on the sheet.",
      keytip: "J",
      largeIcon: <CellStylesIcon />,
      menu: [
        {
          panel: (env, close) => (
            <div className="cell-style-gallery" role="menu">
              {CELL_STYLE_SECTIONS.map((section) => (
                <div key={section}>
                  <div className="color-heading">{section}</div>
                  <div className="cell-style-grid">
                    {CELL_STYLES.filter((s) => s.section === section).map((s) => (
                      <button
                        key={s.name}
                        type="button"
                        role="menuitem"
                        className="cell-style-item"
                        style={{
                          color: s.font?.color,
                          background: s.fill ?? undefined,
                          fontWeight: s.font?.bold ? 700 : undefined,
                          fontStyle: s.font?.italic ? "italic" : undefined,
                          fontSize: s.font?.size ? Math.min(16, s.font.size) : undefined,
                          borderBottom: s.borders?.bottom
                            ? `${s.borders.bottom.style === "double" ? "3px double" : s.borders.bottom.style === "thin" ? "1px solid" : "2px solid"} ${s.borders.bottom.color}`
                            : undefined,
                          outline: s.borders?.top ? `1px solid ${s.borders.top.color}` : undefined,
                          outlineOffset: -1,
                        }}
                        data-keep-light
                        onClick={() => {
                          f(env).apply(cellStylePatch(s), "Cell Style")
                          close()
                        }}
                      >
                        {s.name}
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          ),
        },
        "-",
        "CellStyleNew",
        "CellStylesMerge",
      ],
    },
    { id: "CellStyleNew", label: "New Cell Style..." },
    { id: "CellStylesMerge", label: "Merge Styles..." },

    // ── Cells ──
    {
      id: "InsertCellstMenu",
      label: "Insert",
      description: "Add cells, rows or columns to your sheet.",
      keytip: "I",
      largeIcon: <CellsIcon kind="insert" />,
      run: (env) => env.c.insertOrDelete("insert"),
      menu: ["CellsInsertDialog", "-", "SheetRowsInsert", "SheetColumnsInsert", "-", "SheetInsert"],
    },
    { id: "CellsInsertDialog", label: "Insert Cells...", run: (env) => env.dialog("insert") },
    { id: "SheetRowsInsert", label: "Insert Sheet Rows", run: (env) => env.c.insertLines("rows") },
    { id: "SheetColumnsInsert", label: "Insert Sheet Columns", run: (env) => env.c.insertLines("cols") },
    { id: "SheetInsert", label: "Insert Sheet", shortcut: "Shift+F11", run: (env) => env.c.sheets.insert() },
    {
      id: "TableDeleteRowsAndColumnsMenu",
      label: "Delete",
      description: "Delete cells, rows, columns or sheets from your workbook.",
      keytip: "D",
      largeIcon: <CellsIcon kind="delete" />,
      run: (env) => env.c.insertOrDelete("delete"),
      menu: ["CellsDelete", "-", "SheetRowsDelete", "SheetColumnsDelete", "-", "SheetDelete"],
    },
    { id: "CellsDelete", label: "Delete Cells...", run: (env) => env.dialog("delete") },
    { id: "SheetRowsDelete", label: "Delete Sheet Rows", run: (env) => env.c.deleteLines("rows") },
    { id: "SheetColumnsDelete", label: "Delete Sheet Columns", run: (env) => env.c.deleteLines("cols") },
    { id: "SheetDelete", label: "Delete Sheet", run: (env) => env.c.sheets.remove() },
    {
      id: "FormatCellsMenu",
      label: "Format",
      description: "Change the row height or column width, organize sheets, or protect or hide cells.",
      keytip: "O",
      largeIcon: <CellsIcon kind="format" />,
      menu: [
        { heading: "Cell Size" },
        "RowHeight",
        "RowAutofit",
        "ColumnWidth",
        "ColumnAutofit",
        "ColumnWidthDefault",
        { heading: "Visibility" },
        "HideRows",
        "HideColumns",
        "SheetHide",
        "UnhideRows",
        "UnhideColumns",
        "SheetUnhide",
        { heading: "Organize Sheets" },
        "SheetRename",
        "SheetMoveOrCopy",
        "TabColorPicker",
        { heading: "Protection" },
        "SheetProtect",
        "FormatCellsLocked",
        "FormatCellsDialog",
      ],
    },
    { id: "RowHeight", label: "Row Height...", run: (env) => env.dialog("size:rows") },
    { id: "RowAutofit", label: "AutoFit Row Height", run: (env) => env.dialog("autofit:rows") },
    { id: "ColumnWidth", label: "Column Width...", run: (env) => env.dialog("size:cols") },
    { id: "ColumnAutofit", label: "AutoFit Column Width", run: (env) => env.dialog("autofit:cols") },
    { id: "ColumnWidthDefault", label: "Default Width..." },
    { id: "HideRows", label: "Hide Rows", shortcut: "Ctrl+9", run: (env) => env.c.hideSelected("rows", true) },
    { id: "HideColumns", label: "Hide Columns", shortcut: "Ctrl+0", run: (env) => env.c.hideSelected("cols", true) },
    {
      id: "UnhideRows",
      label: "Unhide Rows",
      shortcut: "Ctrl+Shift+(",
      run: (env) => env.c.hideSelected("rows", false),
    },
    { id: "UnhideColumns", label: "Unhide Columns", run: (env) => env.c.hideSelected("cols", false) },
    { id: "SheetRename", label: "Rename Sheet", run: (env) => env.c.sheets.beginRename() },
    { id: "SheetMoveOrCopy", label: "Move or Copy Sheet...", run: (env) => env.dialog("moveCopySheet") },
    {
      id: "TabColorPicker",
      label: "Tab Color",
      menu: [
        {
          panel: (env, close) => (
            <ColorPanel
              noneLabel="No Color"
              current={env.c.sheet.tabColor}
              onPick={(hex) => {
                env.c.sheets.setColor(env.c.sheetIndex, hex)
                close()
              }}
            />
          ),
        },
      ],
    },
    { id: "SheetHide", label: "Hide Sheet", run: (env) => env.c.sheets.hide() },
    {
      id: "SheetUnhide",
      label: "Unhide Sheet...",
      run: (env) => env.dialog("unhideSheet"),
      enabled: (env) => env.c.sheets.hidden().length > 0,
    },
    {
      id: "SheetProtect",
      label: "Protect Sheet...",
      labelFor: (env) => (env.c.sheets.isProtected() ? "Unprotect Sheet..." : "Protect Sheet..."),
      description: "Prevent unwanted changes by limiting what others can edit on this sheet.",
      run: (env) =>
        env.c.sheets.isProtected()
          ? env.c.sheets.needsPassword()
            ? env.dialog("unprotectSheet")
            : env.c.sheets.unprotect()
          : env.dialog("protectSheet"),
    },
    {
      id: "FormatCellsLocked",
      label: "Lock Cell",
      run: (env) => f(env).apply({ locked: active(env).locked === false ? null : false }, "Lock Cell"),
      pressed: (env) => active(env).locked !== false,
    },
    { id: "FormatCellsDialog", label: "Format Cells...", shortcut: "Ctrl+1", run: (env) => env.dialog("formatCells") },

    // ── Editing ──
    {
      id: "AutoSumMenu",
      label: "AutoSum",
      description: "Automatically add it up. Your total will appear after the selected cells.",
      shortcut: "Alt+=",
      keytip: "U",
      icon: <Sigma {...lucide} />,
      run: (env) => autoSum(env, "SUM"),
      enabled: ready,
      menu: ["AutoSumSum", "AutoSumAverage", "AutoSumCount", "AutoSumMax", "AutoSumMin", "-", "AutoSumMoreFunctions"],
    },
    ...(
      [
        ["AutoSumSum", "Sum", "SUM"],
        ["AutoSumAverage", "Average", "AVERAGE"],
        ["AutoSumCount", "Count Numbers", "COUNT"],
        ["AutoSumMax", "Max", "MAX"],
        ["AutoSumMin", "Min", "MIN"],
      ] as const
    ).map(([id, label, fn]): Command => ({ id, label, run: (env) => autoSum(env, fn), enabled: ready })),
    { id: "AutoSumMoreFunctions", label: "More Functions..." },
    {
      id: "FillMenu",
      label: "Fill",
      description: "Continue a pattern into one or more adjacent cells.",
      keytip: "FI",
      icon: <ArrowDownToLine {...lucide} />,
      menu: [
        "FillDown",
        "FillRight",
        "FillUp",
        "FillLeft",
        "-",
        "FillAcrossWorksheets",
        "FillSeries",
        "FillJustify",
        "QuickCode",
      ],
    },
    { id: "FillDown", label: "Down", shortcut: "Ctrl+D", run: (env) => env.c.fillFromEdge("down") },
    { id: "FillRight", label: "Right", shortcut: "Ctrl+R", run: (env) => env.c.fillFromEdge("right") },
    { id: "FillUp", label: "Up" },
    { id: "FillLeft", label: "Left" },
    { id: "FillAcrossWorksheets", label: "Across Worksheets..." },
    { id: "FillSeries", label: "Series..." },
    { id: "FillJustify", label: "Justify" },
    { id: "QuickCode", label: "Flash Fill", shortcut: "Ctrl+E" },
    {
      id: "ClearMenu",
      label: "Clear",
      description: "Delete everything in the cell, or remove just the formatting, contents, notes or hyperlinks.",
      keytip: "E",
      icon: <Eraser {...lucide} />,
      menu: [
        "ClearAll",
        "ClearFormats",
        "ClearContents",
        "ClearComments",
        "ClearNotes",
        "-",
        "ClearHyperlinks",
        "RemoveHyperlinks",
      ],
    },
    { id: "ClearAll", label: "Clear All", run: (env) => f(env).clearAll() },
    { id: "ClearFormats", label: "Clear Formats", run: (env) => f(env).clearFormats() },
    { id: "ClearContents", label: "Clear Contents", shortcut: "Delete", run: (env) => env.c.clearContents() },
    {
      id: "ClearComments",
      label: "Clear Comments",
      run: (env) => env.c.annotations.clear("comment"),
      enabled: (env) => env.c.annotations.selectionHas("comment"),
    },
    {
      id: "ClearNotes",
      label: "Clear Notes",
      run: (env) => env.c.annotations.clear("note"),
      enabled: (env) => env.c.annotations.selectionHas("note"),
    },
    { id: "ClearHyperlinks", label: "Clear Hyperlinks" },
    { id: "RemoveHyperlinks", label: "Remove Hyperlinks" },
    {
      id: "SortFilterMenu",
      label: "Sort & Filter",
      description: "Organize your data so it's easier to analyze.",
      keytip: "S",
      largeIcon: <ArrowUpDown {...large} />,
      menu: [
        "SortAscendingExcel",
        "SortDescendingExcel",
        "SortCustomExcel",
        "-",
        "Filter",
        "SortClear",
        "FilterReapply",
      ],
    },
    {
      id: "SortAscendingExcel",
      label: "Sort A to Z",
      description: "Lowest to highest.",
      labelFor: (env) => env.c.data.labels().ascending,
      run: (env) => env.c.data.quickSort(false),
    },
    {
      id: "SortDescendingExcel",
      label: "Sort Z to A",
      description: "Highest to lowest.",
      labelFor: (env) => env.c.data.labels().descending,
      run: (env) => env.c.data.quickSort(true),
    },
    {
      id: "SortCustomExcel",
      label: "Custom Sort...",
      description: "Sort data by several columns, by colour, or in a custom order.",
      run: (env) => env.c.data.openSortDialog(),
    },
    {
      id: "SortDialog",
      label: "Sort",
      description: "Sort data by several columns, by colour, or in a custom order.",
      run: (env) => env.c.data.openSortDialog(),
    },
    {
      id: "Filter",
      label: "Filter",
      description:
        "Turn on filtering for the selected cells. Then, click the arrow in the column header to narrow down the data.",
      shortcut: "Ctrl+Shift+L",
      keytip: "T",
      run: (env) => env.c.data.toggleFilter(),
      pressed: (env) => env.c.data.autoFilter !== null,
    },
    {
      id: "SortClear",
      label: "Clear",
      description: "Clear the filter and sort state for the current range of data.",
      keytip: "C",
      run: (env) => env.c.data.clearFilters(),
      enabled: (env) => env.c.data.filtering,
    },
    {
      id: "FilterReapply",
      label: "Reapply",
      description: "Reapply the filter and sort in the current range.",
      keytip: "Y",
      run: (env) => env.c.data.reapply(),
      enabled: (env) => env.c.data.autoFilter !== null,
    },
    {
      id: "SelectMenuExcel",
      label: "Find & Select",
      description: "Find and select specific text, formatting or type of information within the workbook.",
      keytip: "FD",
      largeIcon: <Search {...large} />,
      menu: [
        "FindDialogExcel",
        "ReplaceDialog",
        "GoTo",
        "GoToSpecial",
        "-",
        "SelectFormulas",
        "SelectNotes",
        "SelectConditionalFormatting",
        "SelectConstants",
        "SelectDataValidation",
        "-",
        "SelectObjects",
        "SelectionPane",
      ],
    },
    {
      id: "FindDialogExcel",
      label: "Find...",
      description: "Find text in the sheet or the workbook.",
      shortcut: "Ctrl+F",
      run: (env) => env.c.shell.findReplace?.("find"),
    },
    {
      id: "ReplaceDialog",
      label: "Replace...",
      description: "Search for text and replace it with different text.",
      shortcut: "Ctrl+H",
      run: (env) => env.c.shell.findReplace?.("replace"),
    },
    { id: "GoTo", label: "Go To...", shortcut: "Ctrl+G", run: (env) => env.dialog("goto") },
    {
      id: "GoToSpecial",
      label: "Go To Special...",
      description: "Select cells by what they hold: blanks, formulas, constants, differences and more.",
      run: (env) => env.c.shell.goToSpecial?.(),
    },
    {
      id: "SelectFormulas",
      label: "Formulas",
      run: (env) => env.c.find.special({ kind: "formulas", types: ALL_TYPES }),
    },
    { id: "SelectNotes", label: "Notes", run: (env) => env.c.annotations.selectAll("note") },
    {
      id: "SelectConditionalFormatting",
      label: "Conditional Formatting",
      run: (env) => env.c.find.special({ kind: "conditionalFormats", same: false }),
    },
    {
      id: "SelectConstants",
      label: "Constants",
      run: (env) => env.c.find.special({ kind: "constants", types: ALL_TYPES }),
    },
    {
      id: "SelectDataValidation",
      label: "Data Validation",
      run: (env) => env.c.find.special({ kind: "validation", same: false }),
    },
    { id: "SelectObjects", label: "Select Objects" },
    { id: "SelectionPane", label: "Selection Pane..." },
  ]
  return cmds
}
