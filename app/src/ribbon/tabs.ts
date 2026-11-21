/**
 * The ribbon's tabs as data: groups of controls referring to command ids (see registry.ts). Names, grouping and order
 * follow Excel's ribbon (docs/reference/excel-ribbon.json); button sizes follow Excel at a wide window. Cloud-only
 * groups (Sensitivity, Add-ins store, Copilot, Analyze Data, Data Types, Smart Lookup, Translate) are left out: Cells
 * works offline. In Copilot's place on Home, the AI group holds Cells' opt-in AI tools.
 */

/** A control: `large` is a 32 px button with its label under it; `small` sits in a stack, label beside it. */
export type Control =
  | { kind: "large"; id: string; label?: string }
  | { kind: "small"; id: string; label?: string; /** icon only, also at full size */ iconOnly?: boolean }
  | { kind: "field"; id: string }
  | { kind: "check"; id: string; label?: string }
  | { kind: "label"; text: string }

/** A column of the group: one large control, or a stack of up to three rows. */
export type Column = { kind: "large"; control: Control } | { kind: "stack"; rows: Control[][] }

export type Group = {
  id: string
  label: string
  columns: Column[]
  /** The dialog launcher at the caption's right. */
  launcher?: string
  /** The keytip for the group when it is collapsed to one button (Alt, H, ZF…). */
  keytip?: string
}

export type Tab = {
  id: string
  label: string
  keytip: string
  groups: Group[]
  /** Shown only while its object is selected (Table Design inside a table), after the other tabs. */
  contextual?: boolean
}

const L = (id: string, label?: string): Column => ({ kind: "large", control: { kind: "large", id, label } })
const S = (id: string, label?: string): Control => ({ kind: "small", id, label })
const I = (id: string): Control => ({ kind: "small", id, iconOnly: true })
const F = (id: string): Control => ({ kind: "field", id })
const C = (id: string, label?: string): Control => ({ kind: "check", id, label })
const stack = (...rows: Control[][]): Column => ({ kind: "stack", rows })
/** Small controls stacked one per row. */
const smalls = (...controls: Control[]): Column => stack(...controls.map((c) => [c]))

export const HOME: Tab = {
  id: "TabHome",
  label: "Home",
  keytip: "H",
  groups: [
    {
      id: "GroupClipboard",
      label: "Clipboard",
      launcher: "ShowClipboard",
      keytip: "ZC",
      columns: [L("PasteMenu"), smalls(S("Cut"), S("CopySplitButton"), S("FormatPainter"))],
    },
    {
      id: "GroupFont",
      label: "Font",
      launcher: "FormatCellsFontDialog",
      keytip: "ZF",
      columns: [
        stack(
          [F("Font"), F("FontSize"), I("FontSizeIncrease"), I("FontSizeDecrease")],
          [
            I("Bold"),
            I("Italic"),
            I("UnderlineGallery"),
            I("BordersGallery"),
            I("CellFillColorPicker"),
            I("FontColorPicker"),
          ],
        ),
      ],
    },
    {
      id: "GroupAlignmentExcel",
      label: "Alignment",
      launcher: "CellAlignmentOptions",
      keytip: "ZA",
      columns: [
        stack(
          [I("AlignTopExcel"), I("AlignMiddleExcel"), I("AlignBottomExcel"), I("OrientationMenu")],
          [I("AlignLeft"), I("AlignCenter"), I("AlignRight"), I("IndentDecreaseExcel"), I("IndentIncreaseExcel")],
        ),
        smalls(S("WrapText"), S("MergeCenterMenu")),
      ],
    },
    {
      id: "GroupNumber",
      label: "Number",
      launcher: "FormatCellsNumberDialog",
      keytip: "ZN",
      columns: [
        stack(
          [F("NumberFormatGallery")],
          [I("AccountingFormatMenu"), I("PercentStyle"), I("CommaStyle"), I("DecimalsIncrease"), I("DecimalsDecrease")],
        ),
      ],
    },
    {
      id: "GroupStyles",
      label: "Styles",
      keytip: "ZS",
      columns: [L("ConditionalFormattingMenu"), L("FormatAsTableGallery"), L("CellStylesGallery")],
    },
    {
      id: "GroupCells",
      label: "Cells",
      keytip: "ZL",
      columns: [L("InsertCellstMenu"), L("TableDeleteRowsAndColumnsMenu"), L("FormatCellsMenu")],
    },
    {
      id: "GroupEditingExcel",
      label: "Editing",
      keytip: "ZE",
      columns: [smalls(S("AutoSumMenu"), S("FillMenu"), S("ClearMenu")), L("SortFilterMenu"), L("SelectMenuExcel")],
    },
    // In Copilot's place: the AI the person chose, which may run on this computer.
    { id: "GroupAI", label: "AI", keytip: "ZI", columns: [L("AIToolsMenu")] },
  ],
}

export const INSERT: Tab = {
  id: "TabInsert",
  label: "Insert",
  keytip: "N",
  groups: [
    {
      id: "GroupInsertTablesExcel",
      label: "Tables",
      columns: [L("PivotTableInsert"), L("PivotTableSuggestion"), L("TableInsertExcel")],
    },
    {
      id: "GroupInsertIllustrations",
      label: "Illustrations",
      columns: [
        L("FlyoutAnchorInsertPictures"),
        smalls(S("ShapesInsertGallery"), S("IconInsertFromFile"), S("Insert3DModelDropdown")),
        smalls(S("SmartArtInsert"), S("ScreenshotInsertGallery")),
      ],
    },
    {
      id: "GroupInsertChartsExcel",
      label: "Charts",
      launcher: "ChartInsertGalleryNew",
      columns: [
        L("ChartInsertGalleryNew"),
        stack(
          [
            I("ChartTypeColumnInsertGallery"),
            I("ChartTypeHierarchyInsertGallery"),
            I("ChartTypeStatisticalInsertGallery"),
          ],
          [I("ChartTypeLineInsertGallery"), I("ChartTypeXYScatterInsertGallery")],
          [I("ChartTypePieInsertGallery"), I("ChartTypeOtherInsertGallery")],
        ),
        L("PivotChartInsert", "PivotChart"),
      ],
    },
    {
      id: "GroupSparklinesInsert",
      label: "Sparklines",
      columns: [L("SparklineLineInsert"), L("SparklineColumnInsert"), L("SparklineWinLossInsert")],
    },
    { id: "GroupSlicerInsert", label: "Filters", columns: [L("SlicerInsert", "Slicer"), L("InsertTimeline")] },
    { id: "GroupInsertLinks", label: "Links", columns: [L("InsertLinkGallery")] },
    { id: "GroupInsertComments", label: "Comments", columns: [L("InsertNewComment", "Comment")] },
    {
      id: "GroupInsertText",
      label: "Text",
      columns: [
        L("TextBoxInsertExcel"),
        L("HeaderFooterInsert", "Header & Footer"),
        smalls(S("WordArtInsertGallery"), S("SignatureLineInsertMenu"), S("OleObjectctInsert", "Object")),
      ],
    },
    {
      id: "GroupInsertSymbols",
      label: "Symbols",
      columns: [L("InsertBuildingBlocksEquationsGallery"), L("SymbolInsert", "Symbol")],
    },
  ],
}

/** Draw: pens, a highlighter and the eraser for ink over the sheet (Review › Ink hides or deletes it). */
export const DRAW: Tab = {
  id: "TabDraw",
  label: "Draw",
  keytip: "JI",
  groups: [
    {
      id: "GroupDrawingTools",
      label: "Drawing Tools",
      columns: [L("InkSelect"), L("InkEraser"), L("InkPen0"), L("InkPen1"), L("InkPen2")],
    },
    { id: "GroupDrawTouch", label: "Touch", columns: [L("InkDrawWithTouch")] },
  ],
}

export const PAGE_LAYOUT: Tab = {
  id: "TabPageLayoutExcel",
  label: "Page Layout",
  keytip: "P",
  groups: [
    {
      id: "GroupThemesExcel",
      label: "Themes",
      columns: [L("ThemesGallery"), smalls(S("ThemeColorsGallery"), S("ThemeFontsGallery"), S("ThemeEffectsGallery"))],
    },
    {
      id: "GroupPageSetup",
      label: "Page Setup",
      launcher: "PageSetupPageDialog",
      columns: [
        L("PageMarginsGallery"),
        L("PageOrientationGallery"),
        L("PageSizeGallery"),
        L("PrintAreaMenu"),
        L("PageBreakMenu"),
        L("SheetBackground", "Background"),
        L("PrintTitles"),
      ],
    },
    {
      id: "GroupPageLayoutScaleToFit",
      label: "Scale to Fit",
      launcher: "PageSetupPageDialog",
      columns: [smalls(S("PageScaleToFitWidth"), S("PageScaleToFitHeight"), S("PageScaleToFitScale"))],
    },
    {
      id: "GroupPageLayoutSheetOptions",
      label: "Sheet Options",
      launcher: "PageSetupSheetDialog",
      columns: [
        stack(
          [{ kind: "label", text: "Gridlines" }],
          [C("ViewSheetGridlines", "View")],
          [C("SheetPrintGridlines", "Print")],
        ),
        stack(
          [{ kind: "label", text: "Headings" }],
          [C("ViewSheetHeadings", "View")],
          [C("SheetPrintHeadings", "Print")],
        ),
      ],
    },
    {
      id: "GroupArrange",
      label: "Arrange",
      columns: [
        L("ObjectBringForwardMenu"),
        L("ObjectSendBackwardMenu"),
        L("SelectionPane", "Selection Pane"),
        smalls(S("ObjectAlignMenu"), S("ObjectsGroupMenu"), S("ObjectRotateGallery")),
      ],
    },
  ],
}

export const FORMULAS: Tab = {
  id: "TabFormulas",
  label: "Formulas",
  keytip: "M",
  groups: [
    {
      id: "GroupFunctionLibrary",
      label: "Function Library",
      columns: [
        L("FunctionWizard", "Insert Function"),
        L("AutoSumMenu"),
        L("FunctionsRecentlyUsedtInsertGallery"),
        L("FunctionsFinancialInsertGallery"),
        L("FunctionsLogicalInsertGallery"),
        L("FunctionsTextInsertGallery"),
        L("FunctionsDateTimeInsertGallery"),
        L("FunctionsLookupReferenceInsertGallery"),
        L("FunctionsMathTrigInsertGallery"),
        L("FunctionsMoreInsertGallery"),
      ],
    },
    {
      id: "GroupNamedCells",
      label: "Defined Names",
      columns: [
        L("NameManager"),
        smalls(S("NameDefineMenu"), S("NameUseInFormula"), S("NameCreateFromSelection", "Create from Selection")),
      ],
    },
    {
      id: "GroupFormulaAuditing",
      label: "Formula Auditing",
      columns: [
        smalls(S("TracePrecedents"), S("TraceDependents"), S("TraceRemoveArrowsMenu")),
        smalls(S("ShowFormulas"), S("ErrorCheckingMenu"), S("FormulaEvaluate", "Evaluate Formula")),
        L("WatchWindow"),
      ],
    },
    {
      id: "GroupCalculation",
      label: "Calculation",
      columns: [L("CalculationOptionsMenu"), smalls(S("CalculateNow"), S("CalculateSheet"))],
    },
  ],
}

export const DATA: Tab = {
  id: "TabData",
  label: "Data",
  keytip: "A",
  groups: [
    {
      id: "GroupPowerQueryGetTransformDataSubscription19",
      label: "Get & Transform Data",
      columns: [
        L("PowerQueryGetTransformDataMenuSubscription19"),
        smalls(
          S("GetPowerQueryGetTransformDataFromText"),
          S("GetPowerQueryDataFromWeb"),
          S("GetPowerQueryGetTransformDataFromTable"),
        ),
        smalls(S("PowerQueryRecentSources"), S("GetTransformExistingConnections")),
      ],
    },
    {
      id: "GroupDataQueriesAndConnectionsWorkbookLinks",
      label: "Queries & Connections",
      columns: [
        L("RefreshAllMenu", "Refresh All"),
        smalls(S("ShowQueriesAndConnections"), S("DataRangeProperties"), S("WorkbookLinks")),
      ],
    },
    {
      id: "GroupSortFilter",
      label: "Sort & Filter",
      columns: [
        smalls(S("SortAscendingExcel", ""), S("SortDescendingExcel", "")),
        L("SortDialog", "Sort"),
        L("Filter"),
        smalls(S("SortClear"), S("FilterReapply"), S("AdvancedFilterDialog", "Advanced")),
      ],
    },
    {
      id: "GroupDataTools",
      label: "Data Tools",
      columns: [
        L("ConvertTextToTable", "Text to Columns"),
        stack([I("QuickCode")], [I("RemoveDuplicates")], [I("DataValidationMenu")]),
        stack([I("Consolidate")], [I("AsppRelationships")], [I("DataModelManage")]),
      ],
    },
    {
      id: "GroupForecast",
      label: "Forecast",
      columns: [L("WhatIfAnalysisMenu"), L("ForecastInsert", "Forecast Sheet")],
    },
    {
      id: "GroupOutline",
      label: "Outline",
      launcher: "OutlineSettings",
      columns: [
        L("OutlineGroupMenu"),
        L("OutlineUngroupMenu"),
        L("OutlineSubtotals"),
        smalls(S("OutlineShowDetail"), S("OutlineHideDetail")),
      ],
    },
  ],
}

export const REVIEW: Tab = {
  id: "TabReview",
  label: "Review",
  keytip: "R",
  groups: [
    {
      id: "GroupProofing",
      label: "Proofing",
      columns: [L("Spelling", "Spelling"), smalls(S("Thesaurus", "Thesaurus"), S("StatusWorkbookStatistics"))],
    },
    { id: "GroupAccessibility", label: "Accessibility", columns: [L("MenuAccessibility", "Check Accessibility")] },
    {
      id: "GroupComments",
      label: "Comments",
      columns: [
        L("ReviewNewComment"),
        L("ReviewDeleteComment"),
        smalls(S("ReviewPreviousComment"), S("ReviewNextComment")),
        L("ReviewShowAllComments", "Show Comments"),
      ],
    },
    { id: "GroupCommentsLegacy", label: "Notes", columns: [L("ReviewNoteSplit", "Notes")] },
    {
      id: "GroupProtectExcel",
      label: "Protect",
      columns: [
        L("SheetProtect", "Protect Sheet"),
        L("ReviewRestrictEditing", "Protect Workbook"),
        L("ReviewAllowUsersToEditRangesCoauth"),
        L("SameAsName"),
      ],
    },
    { id: "GroupInk", label: "Ink", columns: [L("HideInkSplitButton", "Hide Ink")] },
  ],
}

export const VIEW: Tab = {
  id: "TabView",
  label: "View",
  keytip: "W",
  groups: [
    {
      id: "GroupWorkbookViews",
      label: "Workbook Views",
      columns: [
        L("ViewNormalViewExcel"),
        L("ViewPageBreakPreviewView"),
        L("ViewPageLayoutView"),
        L("ViewCustomViews", "Custom Views"),
      ],
    },
    {
      id: "GroupViewShowHide",
      label: "Show",
      columns: [
        smalls(C("ViewRulerExcel", "Ruler"), C("GridlinesExcel", "Gridlines")),
        smalls(C("ViewFormulaBar", "Formula Bar"), C("ViewHeadings", "Headings")),
      ],
    },
    {
      id: "GroupZoom",
      label: "Zoom",
      columns: [L("ZoomDialog", "Zoom"), L("ZoomCurrent100"), L("ZoomToSelection")],
    },
    {
      id: "GroupWindow",
      label: "Window",
      columns: [
        L("WindowNew"),
        L("WindowsArrangeAll"),
        L("ViewFreezePanesGallery"),
        stack([S("WindowSplitToggle")], [S("WindowHide")], [S("WindowUnhide", "Unhide")]),
        stack([I("ViewSideBySide")], [I("WindowSideBySideSynchronousScrolling")], [I("WindowResetPosition")]),
        L("WindowSwitchWindowsMenuExcel", "Switch Windows"),
      ],
    },
    { id: "GroupMacros", label: "Macros", columns: [L("MenuMacros")] },
  ],
}

export const HELP: Tab = {
  id: "HelpTab",
  label: "Help",
  keytip: "Y",
  groups: [
    {
      id: "GroupHelpAndSupport",
      label: "Help",
      columns: [
        L("Help"),
        L("ContactUs", "Contact Support"),
        L("OfficeFeedbackHelpTab"),
        L("ShowTraining"),
        L("WhatsNewRecentUpdates"),
      ],
    },
  ],
}

export const TABS: Tab[] = [HOME, INSERT, DRAW, PAGE_LAYOUT, FORMULAS, DATA, REVIEW, VIEW, HELP]

/**
 * Chart Design, while a chart is selected (Excel 365 at a wide window, from memory: no reference screenshot yet). The
 * Format tab isn't built.
 */
export const CHART_DESIGN: Tab = {
  id: "TabChartDesign",
  label: "Chart Design",
  keytip: "JC",
  contextual: true,
  groups: [
    {
      id: "GroupChartLayouts",
      label: "Chart Layouts",
      columns: [L("ChartAddElementMenu"), L("ChartQuickLayoutGallery")],
    },
    { id: "GroupChartStyles", label: "Chart Styles", columns: [L("ChartColorsGallery"), L("ChartStylesGallery")] },
    { id: "GroupChartData", label: "Data", columns: [L("ChartSwitchRowColumn"), L("ChartSelectData")] },
    { id: "GroupChartType", label: "Type", columns: [L("ChartChangeType")] },
    { id: "GroupChartLocation", label: "Location", columns: [L("ChartMove")] },
    { id: "GroupChartAccessibility", label: "Accessibility", columns: [L("ObjectAlternativeText")] },
  ],
}

/** Arrange, as Shape Format, Picture Format and Page Layout show it. */
const ARRANGE: Group = {
  id: "GroupArrange",
  label: "Arrange",
  columns: [
    L("ObjectBringForwardMenu"),
    L("ObjectSendBackwardMenu"),
    L("SelectionPane", "Selection Pane"),
    smalls(S("ObjectAlignMenu"), S("ObjectsGroupMenu"), S("ObjectRotateGallery")),
  ],
}

const SIZE: Group = { id: "GroupSize", label: "Size", columns: [smalls(F("ShapeHeight"), F("ShapeWidth"))] }

/** Equation, while an equation is selected (beside Shape Format, as in Excel). */
export const EQUATION: Tab = {
  id: "TabEquationToolsDesign",
  label: "Equation",
  keytip: "JE",
  contextual: true,
  groups: [
    { id: "GroupEquationTools", label: "Tools", columns: [L("InsertBuildingBlocksEquationsGallery", "Equation")] },
    {
      id: "GroupEquationConversions",
      label: "Conversions",
      columns: [smalls(S("EquationProfessional"), S("EquationLinearFormat"))],
    },
    { id: "GroupEquationSymbols", label: "Symbols", columns: [L("EquationSymbolsGallery", "Symbols")] },
    {
      id: "GroupEquationStructures",
      label: "Structures",
      columns: [
        L("EquationFractionGallery", "Fraction"),
        L("EquationScriptGallery", "Script"),
        L("EquationRadicalGallery", "Radical"),
        L("EquationIntegralGallery", "Integral"),
        L("EquationLargeOperatorGallery", "Large Operator"),
        L("EquationBracketGallery", "Bracket"),
        L("EquationFunctionGallery", "Function"),
        L("EquationAccentGallery", "Accent"),
        L("EquationLimitLogGallery", "Limit and Log"),
        L("EquationOperatorGallery", "Operator"),
        L("EquationMatrixGallery", "Matrix"),
      ],
    },
  ],
}

/** Shape Format, while a shape, text box or WordArt is selected. */
export const SHAPE_FORMAT: Tab = {
  id: "TabDrawingToolsFormat",
  label: "Shape Format",
  keytip: "JD",
  contextual: true,
  groups: [
    {
      id: "GroupInsertShapes",
      label: "Insert Shapes",
      columns: [smalls(S("ShapesInsertGallery"), S("TextBoxInsertExcel"))],
    },
    {
      id: "GroupShapeStyles",
      label: "Shape Styles",
      columns: [
        L("ShapeStylesGallery"),
        smalls(S("ShapeFillColorPicker"), S("ShapeOutlineColorPicker"), S("ShapeEffectsMenu")),
      ],
    },
    {
      id: "GroupWordArtStyles",
      label: "WordArt Styles",
      columns: [smalls(S("TextFillColorPicker"), S("TextOutlineColorPicker"), S("TextEffectsMenu"))],
    },
    { id: "GroupShapeAccessibility", label: "Accessibility", columns: [L("ObjectAlternativeText")] },
    ARRANGE,
    SIZE,
  ],
}

/** SmartArt Design, while a SmartArt is selected (Excel's, without its styles gallery: Cells draws Office's flat style). */
export const SMARTART_DESIGN: Tab = {
  id: "TabSmartArtDesign",
  label: "SmartArt Design",
  keytip: "JS",
  contextual: true,
  groups: [
    {
      id: "GroupSmartArtCreateGraphic",
      label: "Create Graphic",
      columns: [
        L("SmartArtAddShape", "Add Shape"),
        smalls(S("SmartArtPromote"), S("SmartArtDemote"), S("SmartArtTextPane")),
        smalls(S("SmartArtMoveUp"), S("SmartArtMoveDown")),
      ],
    },
    { id: "GroupSmartArtLayouts", label: "Layouts", columns: [L("SmartArtChangeLayoutGallery", "Change Layout")] },
    {
      id: "GroupSmartArtStyles",
      label: "SmartArt Styles",
      columns: [L("SmartArtChangeColorsGallery", "Change Colors")],
    },
    {
      id: "GroupSmartArtReset",
      label: "Reset",
      columns: [L("SmartArtResetGraphic", "Reset Graphic"), L("SmartArtConvertToShapes", "Convert")],
    },
  ],
}

/** Format, while a SmartArt is selected: its frame's place, order and size. */
export const SMARTART_FORMAT: Tab = {
  id: "TabSmartArtFormat",
  label: "Format",
  keytip: "JO",
  contextual: true,
  groups: [
    { id: "GroupSmartArtAccessibility", label: "Accessibility", columns: [L("ObjectAlternativeText")] },
    ARRANGE,
    SIZE,
  ],
}

/** Picture Format, while a picture or an icon is selected. */
export const PICTURE_FORMAT: Tab = {
  id: "TabPictureToolsFormat",
  label: "Picture Format",
  keytip: "JP",
  contextual: true,
  groups: [
    { id: "GroupPictureAdjust", label: "Adjust", columns: [smalls(S("PictureReset"))] },
    { id: "GroupPictureStyles", label: "Picture Styles", columns: [smalls(S("PictureBorderColorPicker"))] },
    { id: "GroupPictureAccessibility", label: "Accessibility", columns: [L("ObjectAlternativeText")] },
    ARRANGE,
    SIZE,
  ],
}

/** 3D Model, while a 3D model is selected (Excel's 3D Model tab, with its turns as fields beside the views). */
export const MODEL3D_FORMAT: Tab = {
  id: "TabModel3DFormat",
  label: "3D Model",
  keytip: "J3",
  contextual: true,
  groups: [
    {
      id: "GroupModel3DAdjust",
      label: "Adjust",
      columns: [L("Model3DChange", "Change 3D Model"), L("Model3DReset", "Reset 3D Model")],
    },
    {
      id: "GroupModel3DViews",
      label: "3D Model Views",
      columns: [
        L("Model3DViewsGallery", "Views"),
        stack(
          [{ kind: "label", text: "X" }, F("Model3DRotationX")],
          [{ kind: "label", text: "Y" }, F("Model3DRotationY")],
          [{ kind: "label", text: "Z" }, F("Model3DRotationZ")],
        ),
      ],
    },
    { id: "GroupModel3DAccessibility", label: "Accessibility", columns: [L("ObjectAlternativeText")] },
    {
      id: "GroupArrange",
      label: "Arrange",
      columns: [
        L("ObjectBringForwardMenu"),
        L("ObjectSendBackwardMenu"),
        L("SelectionPane", "Selection Pane"),
        smalls(S("ObjectAlignMenu")),
      ],
    },
    SIZE,
  ],
}

/** Sparkline, while the active cell holds a sparkline (Excel 365's tab at a wide window). */
export const SPARKLINE: Tab = {
  id: "TabSparkline",
  label: "Sparkline",
  keytip: "JS",
  contextual: true,
  groups: [
    { id: "GroupSparklineData", label: "Sparkline", columns: [L("SparklineEditData")] },
    {
      id: "GroupSparklineType",
      label: "Type",
      columns: [L("SparklineTypeLine"), L("SparklineTypeColumn"), L("SparklineTypeWinLoss")],
    },
    {
      id: "GroupSparklineShow",
      label: "Show",
      columns: [
        smalls(C("SparklineShowHigh"), C("SparklineShowLow"), C("SparklineShowNegative")),
        smalls(C("SparklineShowFirst"), C("SparklineShowLast"), C("SparklineShowMarkers")),
      ],
    },
    {
      id: "GroupSparklineStyle",
      label: "Style",
      columns: [L("SparklineStyleGallery"), smalls(S("SparklineColor"), S("SparklineMarkerColor"))],
    },
    {
      id: "GroupSparklineGroup",
      label: "Group",
      columns: [L("SparklineAxis"), smalls(S("SparklineGroup"), S("SparklineUngroup"), S("SparklineClear"))],
    },
  ],
}

/** Slicer, while slicers are selected (Excel's Slicer tab). */
export const SLICER: Tab = {
  id: "TabSlicer",
  label: "Slicer",
  keytip: "JI",
  contextual: true,
  groups: [
    {
      id: "GroupSlicerProperties",
      label: "Slicer",
      columns: [
        smalls({ kind: "label", text: "Slicer Caption:" }, F("SlicerCaption"), S("SlicerSettings")),
        L("SlicerReportConnections"),
      ],
    },
    { id: "GroupSlicerStyles", label: "Slicer Styles", columns: [L("SlicerStyleGallery")] },
    ARRANGE,
    {
      id: "GroupSlicerButtons",
      label: "Buttons",
      columns: [
        stack(
          [{ kind: "label", text: "Columns:" }, F("SlicerColumns")],
          [{ kind: "label", text: "Height:" }, F("SlicerButtonHeight")],
          [{ kind: "label", text: "Width:" }, F("SlicerButtonWidth")],
        ),
      ],
    },
    SIZE,
  ],
}

/** Timeline, while timelines are selected (Excel 365's: Timeline, Timeline Styles, Arrange, Size, Show). */
export const TIMELINE: Tab = {
  id: "TabTimeline",
  label: "Timeline",
  keytip: "JL",
  contextual: true,
  groups: [
    {
      id: "GroupTimelineProperties",
      label: "Timeline",
      columns: [
        smalls({ kind: "label", text: "Timeline Caption:" }, F("TimelineCaption")),
        L("TimelineReportConnections"),
      ],
    },
    { id: "GroupTimelineStyles", label: "Timeline Styles", columns: [L("TimelineStyleGallery")] },
    ARRANGE,
    SIZE,
    {
      id: "GroupTimelineShow",
      label: "Show",
      columns: [
        smalls(C("TimelineShowHeader"), C("TimelineShowSelectionLabel")),
        smalls(C("TimelineShowScrollbar"), C("TimelineShowTimeLevel")),
      ],
    },
  ],
}

/**
 * Table Design, while the active cell is in a table. Built from Excel 365 at a wide window (no reference screenshot
 * in docs/reference yet). External Table Data (Export, Refresh, Unlink…) is left out: Cells has no linked tables.
 */
export const TABLE_DESIGN: Tab = {
  id: "TabTableDesign",
  label: "Table Design",
  keytip: "JT",
  contextual: true,
  groups: [
    {
      id: "GroupTableProperties",
      label: "Properties",
      columns: [smalls({ kind: "label", text: "Table Name:" }, F("TableName"), S("TableResize"))],
    },
    {
      id: "GroupTableTools",
      label: "Tools",
      columns: [
        smalls(S("TableSummarizeWithPivot"), S("TableRemoveDuplicates"), S("TableConvertToRange")),
        L("TableInsertSlicer"),
      ],
    },
    {
      id: "GroupTableStyleOptions",
      label: "Table Style Options",
      columns: [
        smalls(C("TableStyleHeaderRow"), C("TableStyleTotalRow"), C("TableStyleBandedRows")),
        smalls(C("TableStyleFirstColumn"), C("TableStyleLastColumn"), C("TableStyleBandedColumns")),
        smalls(C("TableStyleFilterButton")),
      ],
    },
    {
      id: "GroupTableStyles",
      label: "Table Styles",
      columns: [L("TableStylesGallery")],
    },
  ],
}

/**
 * PivotTable Analyze, while the active cell is in a PivotTable. As Excel 365 at a wide window, with what isn't built
 * yet (grouping, slicers, calculated fields, PivotCharts) greyed out; OLAP Tools and Relationships (Data Model only)
 * are left out.
 */
export const PIVOT_ANALYZE: Tab = {
  id: "TabPivotTableAnalyze",
  label: "PivotTable Analyze",
  keytip: "JT",
  contextual: true,
  groups: [
    {
      id: "GroupPivotTable",
      label: "PivotTable",
      columns: [smalls({ kind: "label", text: "PivotTable Name:" }, F("PivotTableName"), S("PivotTableOptionsMenu"))],
    },
    {
      id: "GroupPivotActiveField",
      label: "Active Field",
      columns: [
        smalls({ kind: "label", text: "Active Field:" }, F("PivotActiveField"), S("PivotFieldSettings")),
        L("PivotDrillDown"),
        L("PivotDrillUp"),
        smalls(S("PivotExpandField"), S("PivotCollapseField")),
      ],
    },
    {
      id: "GroupPivotGroup",
      label: "Group",
      columns: [smalls(S("PivotGroupSelection"), S("PivotUngroup"), S("PivotGroupField"))],
    },
    {
      id: "GroupPivotFilter",
      label: "Filter",
      columns: [smalls(S("PivotInsertSlicer"), S("PivotInsertTimeline"), S("PivotFilterConnections"))],
    },
    {
      id: "GroupPivotData",
      label: "Data",
      columns: [L("PivotRefreshMenu"), L("PivotChangeDataSource")],
    },
    {
      id: "GroupPivotActions",
      label: "Actions",
      columns: [smalls(S("PivotClearMenu"), S("PivotSelectMenu"), S("PivotMove"))],
    },
    {
      id: "GroupPivotCalculations",
      label: "Calculations",
      columns: [L("PivotCalculatedFields")],
    },
    {
      id: "GroupPivotTools",
      label: "Tools",
      columns: [L("PivotChartInsert", "PivotChart"), L("PivotTableSuggestion")],
    },
    {
      id: "GroupPivotShow",
      label: "Show",
      columns: [L("PivotFieldList"), L("PivotButtons"), L("PivotFieldHeaders")],
    },
  ],
}

/**
 * PivotChart Analyze, while a PivotChart is selected (before Chart Design, as in Excel 365): refreshing its
 * PivotTable and its field buttons. Excel's PivotChart name box, Active Field, slicers and timelines, Change Data
 * Source, Clear and Move Chart aren't built for PivotCharts yet.
 */
export const PIVOT_CHART_ANALYZE: Tab = {
  id: "TabPivotChartAnalyze",
  label: "PivotChart Analyze",
  keytip: "JA",
  contextual: true,
  groups: [
    { id: "GroupPivotChartData", label: "Data", columns: [L("PivotChartRefresh")] },
    { id: "GroupPivotChartShowHide", label: "Show/Hide", columns: [L("PivotChartFieldButtons")] },
  ],
}

/** Design (PivotTable Tools), while the active cell is in a PivotTable. */
export const PIVOT_DESIGN: Tab = {
  id: "TabPivotTableDesign",
  label: "Design",
  keytip: "JY",
  contextual: true,
  groups: [
    {
      id: "GroupPivotLayout",
      label: "Layout",
      columns: [
        L("PivotSubtotalsMenu"),
        L("PivotGrandTotalsMenu"),
        L("PivotReportLayoutMenu"),
        L("PivotBlankRowsMenu"),
      ],
    },
    {
      id: "GroupPivotStyleOptions",
      label: "PivotTable Style Options",
      columns: [
        smalls(C("PivotStyleRowHeaders"), C("PivotStyleColumnHeaders")),
        smalls(C("PivotStyleBandedRows"), C("PivotStyleBandedColumns")),
      ],
    },
    {
      id: "GroupPivotStyles",
      label: "PivotTable Styles",
      columns: [L("PivotStylesGallery")],
    },
  ],
}

/** Every command id a tab shows, in order (for KeyTips, tests and Tell Me). */
export function controlIds(tab: Tab): string[] {
  const out: string[] = []
  for (const g of tab.groups) {
    for (const col of g.columns) {
      const controls = col.kind === "large" ? [col.control] : col.rows.flat()
      for (const c of controls) if (c.kind !== "label") out.push(c.id)
    }
    if (g.launcher) out.push(g.launcher)
  }
  return out
}
