import { sparklineCommands } from "./sparkline-commands"
import { linkCommands } from "./link-commands"
import { pageLayoutCommands } from "./page-layout-commands"
import { themeCommands } from "./theme-commands"
import { insertObjectCommands } from "./insert-object-commands"
import { windowCommands } from "./window-commands"
import { equationCommands } from "./equation-commands"
import { model3dCommands } from "./model3d-commands"
import { smartArtCommands } from "./smartart-commands"
import { Calculator, CircleHelp, Columns2, Eye, Scan, Snowflake, Sparkles, Table2, ZoomIn } from "lucide-react"
import labels from "./excel-labels.json"
import { conditionalCommands } from "./cf-commands"
import { homeCommands } from "./home-commands"
import { auditCommands } from "./audit-commands"
import { outlineCommands } from "./outline-commands"
import { whatIfCommands } from "./whatif-commands"
import { dataToolCommands } from "./datatools-commands"
import { formulaCommands, nameCommand } from "./formula-commands"
import { pictureFor } from "./pictures"
import { qatCommands } from "./qat-commands"
import { reviewCommands } from "./review-commands"
import { CommandRegistry, type Command, type RibbonEnv } from "./registry"
import { CHART_DESIGN, EQUATION, PICTURE_FORMAT, SHAPE_FORMAT, SPARKLINE, TABLE_DESIGN, TABS, controlIds } from "./tabs"
import { chartCommands } from "./chart-commands"
import { objectCommands } from "./object-commands"
import { tableCommands } from "./table-commands"

const large = { size: 32, strokeWidth: 1.25 }
const small = { size: 16, strokeWidth: 1.6 }

const EXCEL_LABELS = labels as Record<string, string>

/** Excel's name for a control id, from the reference. */
export function excelLabel(id: string): string | undefined {
  return EXCEL_LABELS[id]
}

/** View › 100%: the figure itself, as Excel's button shows it. */
function Percent100Icon() {
  return (
    <svg width={32} height={32} viewBox="0 0 32 32" aria-hidden>
      <rect x={3.5} y={6.5} width={25} height={19} rx={2} fill="none" stroke="currentColor" strokeWidth={1.25} />
      <text x={16} y={20} textAnchor="middle" fontSize={9} fontWeight={600} fill="currentColor">
        100
      </text>
    </svg>
  )
}

/** Names the reference lacks (controls Excel labels only in some editions). */
const LABELS: Record<string, string> = {
  FunctionsMoreInsertGallery: "More Functions",
  InsertNewComment: "Comment",
  ReviewNoteSplit: "Notes",
  ChartTypeOtherInsertGallery: "Insert Waterfall, Funnel, Stock, Surface or Radar Chart",
}

/** Commands on the other tabs that already work. */
function otherCommands(): Command[] {
  const fromText: Command = {
    id: "GetPowerQueryGetTransformDataFromText",
    label: "From Text/CSV",
    description: "Import data from a text or CSV file onto a new sheet, choosing how its columns split.",
    run: (env) => void env.files?.importText(),
    enabled: (env) => Boolean(env.files),
  }
  const ruler: Command = {
    id: "ViewRulerExcel",
    label: "Ruler",
    description: "Show the rulers next to the document, to measure and line things up. They show in Page Layout view.",
    keytip: "VR",
    check: true,
    enabled: (env) => env.c.view === "pageLayout",
    run: (env) => {
      env.c.showRuler = !env.c.showRuler
      env.c.changed()
    },
    pressed: (env) => env.c.showRuler,
  }
  const gridlines: Command = {
    id: "GridlinesExcel",
    label: "Gridlines",
    description: "Show the lines between rows and columns in the sheet to make it easier to read.",
    keytip: "VG",
    check: true,
    run: (env) => {
      env.c.showGridlines = !env.c.showGridlines
      env.c.changed()
    },
    pressed: (env) => env.c.showGridlines,
  }
  const removeDuplicates: Command = {
    id: "RemoveDuplicates",
    label: "Remove Duplicates",
    description:
      "Delete duplicate rows from a sheet. You can choose which columns should be checked for duplicate information.",
    keytip: "M",
    run: (env) => env.c.data.openRemoveDuplicates(),
  }
  const textToColumns: Command = {
    id: "ConvertTextToTable",
    label: "Text to Columns",
    description:
      "Split a single column of text into multiple columns. For example, you can separate a column of full names into separate first and last name columns.",
    keytip: "E",
    run: (env) => env.c.data.openTextToColumns(),
  }
  const unprotected = (env: RibbonEnv) => !env.c.sheet.protection
  const dataValidation: Command[] = [
    {
      id: "DataValidationMenu",
      label: "Data Validation",
      description:
        "Pick from a list of rules to limit the type of data that can be entered in a cell. For example, you can provide a list of values like 1, 2, and 3 or only allow numbers greater than 1000 as valid entries.",
      keytip: "V",
      run: (env) => env.c.validation.open(),
      enabled: unprotected,
      menu: ["DataValidation", "DataValidationCircleInvalid", "DataValidationClearValidationCircles"],
    },
    {
      id: "DataValidation",
      label: "Data Validation...",
      keytip: "V",
      run: (env) => env.c.validation.open(),
      enabled: unprotected,
    },
    {
      id: "DataValidationCircleInvalid",
      label: "Circle Invalid Data",
      keytip: "I",
      run: (env) => env.c.validation.circleInvalid(),
    },
    {
      id: "DataValidationClearValidationCircles",
      label: "Clear Validation Circles",
      keytip: "R",
      run: (env) => env.c.validation.clearCircles(),
    },
  ]
  return [
    ...dataValidation,
    fromText,
    removeDuplicates,
    textToColumns,
    gridlines,
    ruler,
    { ...gridlines, id: "ViewSheetGridlines", label: "View", keytip: "VG" },
    {
      id: "ViewHeadings",
      label: "Headings",
      description: "Show row and column headings.",
      keytip: "VH",
      check: true,
      run: (env) => {
        env.c.showHeadings = !env.c.showHeadings
        env.c.changed()
      },
      pressed: (env) => env.c.showHeadings,
    },
    {
      id: "ViewSheetHeadings",
      label: "View",
      check: true,
      run: (env) => {
        env.c.showHeadings = !env.c.showHeadings
        env.c.changed()
      },
      pressed: (env) => env.c.showHeadings,
    },
    {
      id: "ViewFormulaBar",
      label: "Formula Bar",
      description: "Show the Formula Bar, where you can type text and formulas into cells.",
      keytip: "VF",
      check: true,
      run: (env) => env.dialog("toggle:formulaBar"),
      pressed: (env) => env.c.showFormulaBar,
    },
    {
      id: "ViewNormalViewExcel",
      label: "Normal",
      description: "View the document normally.",
      keytip: "L",
      largeIcon: <Table2 {...large} />,
      run: (env) => env.c.setView("normal"),
      pressed: (env) => env.c.view === "normal",
    },
    {
      id: "ViewCustomViews",
      label: "Custom Views",
      description:
        "Save the current display and print settings as a custom view that you can quickly apply in the future.",
      keytip: "C",
      run: (env) => env.dialog("customViews"),
    },
    {
      id: "ViewPageLayoutView",
      label: "Page Layout",
      description:
        "See what the printed pages will look like: where each page starts and ends, and the header and footer.",
      keytip: "P",
      run: (env) => env.c.setView("pageLayout"),
      pressed: (env) => env.c.view === "pageLayout",
    },
    {
      id: "ViewPageBreakPreviewView",
      label: "Page Break Preview",
      description:
        "See where the pages break when the sheet prints. Drag a blue line to move a page break, or the edge of what prints.",
      keytip: "I",
      run: (env) => env.c.setView("pageBreakPreview"),
      pressed: (env) => env.c.view === "pageBreakPreview",
    },
    {
      id: "ViewFreezePanesGallery",
      label: "Freeze Panes",
      description: "Freeze a portion of the sheet to keep it visible while you scroll through the rest of the sheet.",
      keytip: "F",
      largeIcon: <Snowflake {...large} />,
      // Page Layout has no frozen panes (Excel greys it out there).
      enabled: (env) => env.c.view !== "pageLayout",
      menu: (env) => [
        env.c.freeze ? "ViewUnfreezePanes" : "ViewFreezePanes",
        "ViewFreezeTopRow",
        "ViewFreezeFirstColumn",
      ],
    },
    {
      id: "WindowSplitToggle",
      label: "Split",
      description:
        "Divide the window into different panes that each scroll separately. Click Split again, or double-click a split bar, to remove it.",
      keytip: "S",
      icon: <Columns2 {...small} />,
      run: (env) => env.c.toggleSplit(),
      enabled: (env) => env.c.view !== "pageLayout",
      pressed: (env) => env.c.split !== null,
    },
    {
      id: "ViewFreezePanes",
      label: "Freeze Panes",
      description:
        "Keep rows and columns visible while the rest of the worksheet scrolls (based on current selection).",
      run: (env) => env.c.freezePanes(),
    },
    {
      id: "ViewUnfreezePanes",
      label: "Unfreeze Panes",
      description: "Unlock all rows and columns to scroll through the entire worksheet.",
      run: (env) => env.c.unfreeze(),
    },
    {
      id: "ViewFreezeTopRow",
      label: "Freeze Top Row",
      description: "Keep the top row visible while scrolling through the rest of the worksheet.",
      run: (env) => env.c.freezeTopRow(),
    },
    {
      id: "ViewFreezeFirstColumn",
      label: "Freeze First Column",
      description: "Keep the first column visible while scrolling through the rest of the worksheet.",
      run: (env) => env.c.freezeFirstColumn(),
    },
    {
      id: "ZoomDialog",
      label: "Zoom",
      description: "Zoom to the level that works for you.",
      keytip: "Q",
      largeIcon: <ZoomIn {...large} />,
      run: (env) => env.dialog("zoom"),
    },
    {
      id: "ZoomCurrent100",
      label: "100%",
      description: "Zoom the document to 100% of the normal size.",
      keytip: "J",
      largeIcon: <Percent100Icon />,
      run: (env) => env.c.setZoom(1),
    },
    {
      id: "ZoomToSelection",
      label: "Zoom to Selection",
      description: "Zoom the workbook so that the currently selected range of cells fills the entire window.",
      keytip: "G",
      largeIcon: <Scan {...large} />,
      run: (env) => env.dialog("zoomToSelection"),
    },
    {
      id: "ShowFormulas",
      label: "Show Formulas",
      description: "Display the formula in each cell instead of the resulting value.",
      shortcut: "Ctrl+`",
      keytip: "H",
      icon: <Eye {...small} />,
      run: (env) => env.c.toggleShowFormulas(),
      pressed: (env) => env.c.showFormulas,
    },
    {
      id: "CalculationOptionsMenu",
      label: "Calculation Options",
      description: "Choose to calculate formulas automatically or manually.",
      keytip: "X",
      largeIcon: <Calculator {...large} />,
      menu: ["CalculationAutomatic", "CalculationManual"],
    },
    {
      id: "CalculationAutomatic",
      label: "Automatic",
      run: (env) => env.c.wb.setCalcMode("automatic"),
      pressed: (env) => env.c.wb.calcStatus.mode === "automatic",
    },
    {
      id: "CalculationManual",
      label: "Manual",
      run: (env) => env.c.wb.setCalcMode("manual"),
      pressed: (env) => env.c.wb.calcStatus.mode === "manual",
    },
    {
      id: "CalculateNow",
      label: "Calculate Now",
      description: "Calculate the entire workbook now.",
      shortcut: "F9",
      keytip: "B",
      icon: <Calculator {...small} />,
      run: (env) => env.c.wb.calculate(false),
    },
    {
      id: "CalculateSheet",
      label: "Calculate Sheet",
      description: "Calculate the current sheet now.",
      shortcut: "Shift+F9",
      keytip: "J",
      icon: <Calculator {...small} />,
      run: (env) => env.c.wb.calculate(false),
    },
    {
      id: "Help",
      label: "Help",
      description: "Get help using Cells.",
      shortcut: "F1",
      keytip: "H",
      largeIcon: <CircleHelp {...large} />,
    },
    {
      id: "WhatsNewRecentUpdates",
      label: "What's New",
      description: "See what changed in the latest version of Cells.",
      keytip: "N",
      largeIcon: <Sparkles {...large} />,
      run: (env) => env.backstage("about"),
    },
  ]
}

/** The registry: built commands, plus every other control on the tabs, by Excel's name, shown as not built yet. */
/** Every command Cells defines, from each tab's file. An id appears once: a second one would replace the first. */
export function allCommands(): Command[] {
  return [
    ...homeCommands(),
    ...conditionalCommands(),
    ...reviewCommands(),
    ...formulaCommands(),
    ...auditCommands(),
    ...outlineCommands(),
    ...whatIfCommands(),
    ...dataToolCommands(),
    ...qatCommands(),
    ...tableCommands(),
    ...chartCommands(),
    ...objectCommands(),
    ...pageLayoutCommands(),
    ...themeCommands(),
    ...windowCommands(),
    ...insertObjectCommands(),
    ...equationCommands(),
    ...model3dCommands(),
    ...smartArtCommands(),
    ...sparklineCommands(),
    ...linkCommands(),
    ...otherCommands(),
  ]
}

export function buildRegistry(): CommandRegistry {
  const registry = new CommandRegistry().add(...allCommands())
  for (const tab of [...TABS, TABLE_DESIGN, CHART_DESIGN, SHAPE_FORMAT, PICTURE_FORMAT, EQUATION, SPARKLINE]) {
    for (const id of controlIds(tab)) {
      if (!registry.get(id)) registry.add({ id, label: LABELS[id] ?? excelLabel(id) ?? id })
    }
  }
  registry.fallback = (id) => (id.startsWith("Name:") ? nameCommand(id.slice(5)) : undefined)
  // Pictures for commands that have none of their own.
  for (const command of registry.all()) {
    if (!command.icon) command.icon = pictureFor(command.id, 16)
    if (!command.largeIcon) command.largeIcon = pictureFor(command.id, 32)
  }
  return registry
}
