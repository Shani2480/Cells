import { PageSetupDialog, type PageSetupTab } from "@/shell/PageSetupDialog"
import {
  isChart,
  isEquation,
  PIVOT_FN_LABEL,
  SHOW_AS_BASE_FIELD,
  SHOW_AS_CHOICES,
  type PivotFn,
  type PivotShowAs,
  parseCell,
  rangeName,
  regionalFrom,
  setRegional,
  type Range,
  type SheetSlicer,
  type SheetTimeline,
} from "@cells/core"
import {
  lazy,
  Suspense,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from "react"
import { columnChars, charsToPixels, pointsToPixels, rowPoints } from "@/grid/autofit"
import { WindowsArea } from "@/shell/WindowsArea"
import { sheetEditor } from "@/ribbon/controls"
import { GridView, type GridContextMenu } from "@/grid/GridView"
import { usePlatform } from "@/platform/context"
import type { SettingsPort } from "@cells/platform"
import { loadSettings, useSettings } from "@/settings/settings"
import { AiDialog } from "@/ai/AiDialog"
import type { AiTool } from "@/ai/prompts"
import { OptionsDialog, type OptionsPage } from "@/settings/OptionsDialog"
import { FirstRunSetup } from "@/settings/FirstRunSetup"
import { parseOpenRequest } from "@/links/link"
import { LiveLinks } from "@/links/live"
import { SheetController } from "@/sheet/controller"
import { absoluteName } from "@/sheet/text-columns"
import { buildRegistry } from "@/ribbon/commands"
import { pictureFor } from "@/ribbon/pictures"
import type { RibbonEnv } from "@/ribbon/registry"
import { ActionSheet } from "@/phone/ActionSheet"
import { PhoneShell } from "@/phone/PhoneShell"
import { installBackHandler, syncSystemBars } from "@/platform/android-files"
import { useShellKind } from "@/phone/shell-choice"
import { Ribbon, type RibbonHandle } from "@/ribbon/Ribbon"
import {
  CHART_DESIGN,
  PIVOT_CHART_ANALYZE,
  EQUATION,
  HELP,
  MODEL3D_FORMAT,
  SMARTART_DESIGN,
  SMARTART_FORMAT,
  PICTURE_FORMAT,
  SHAPE_FORMAT,
  SLICER,
  TIMELINE,
  SPARKLINE,
  TABLE_DESIGN,
  PIVOT_ANALYZE,
  PIVOT_DESIGN,
  TABS,
} from "@/ribbon/tabs"
import {
  SparklineAxisDialog,
  SparklineDataDialog,
  SparklineDateDialog,
  SparklineEmptyDialog,
} from "@/shell/SparklineDialogs"
import {
  InsertSlicersDialog,
  InsertTimelinesDialog,
  SlicerConnectionsDialog,
  SlicerSettingsDialog,
} from "@/shell/SlicerDialogs"
import type { SparklineDialogInit } from "@/sheet/sparklines"
import { Backstage } from "@/shell/Backstage"
import { isTextField, menuBarAction } from "@/shell/menu-bar"
import { whatsNewOnStart } from "@/about/release-notes"
import { Dialog } from "@/shell/Dialog"
import { RemoveDuplicatesDialog } from "@/shell/RemoveDuplicatesDialog"
import { DataValidationDialog, ValidationAlertDialog } from "@/shell/DataValidationDialog"
import type { AlertAnswer, ValidationAlert, ValidationDialogInit } from "@/sheet/validation"
import { FindReplaceDialog, GoToSpecialDialog } from "@/shell/FindDialogs"
import { CustomFilterDialog, Top10Dialog } from "@/shell/FilterDialogs"
import { PivotFilterDialog } from "@/shell/PivotFilterDialogs"
import { RecommendedPivotsDialog } from "@/shell/RecommendedPivotsDialog"
import { CreatePivotDialog } from "@/shell/CreatePivotDialog"
import { PivotFieldsPane } from "@/shell/PivotFieldsPane"
import {
  FieldSettingsDialog,
  PivotOptionsDialog,
  ShowValuesAsDialog,
  CalculatedFieldDialog,
  GroupingDialog,
  ValueFieldSettingsDialog,
} from "@/shell/PivotSettingsDialogs"
import { Check } from "lucide-react"
import type { PivotFilterRequest } from "@/sheet/pivots"
import type { FilterDialogRequest } from "@/sheet/data"
import { SortDialog, SortWarningDialog } from "@/shell/SortDialog"
import { FormatCellsDialog, type FormatTab } from "@/shell/format-cells/FormatCellsDialog"
import { TooltipProvider } from "@/ui/tooltip"
import { ClipboardPane } from "@/shell/ClipboardPane"
import { CommentsPane } from "@/shell/CommentsPane"
import { AccessibilityPane } from "@/shell/AccessibilityPane"
import { SpellingDialog } from "@/shell/SpellingDialog"
import { SelectionPane } from "@/shell/SelectionPane"
import { SmartArtTextPane } from "@/shell/SmartArtTextPane"
// The icon set is large: it loads the first time Insert › Icons opens.
const SmartArtDialog = lazy(() => import("@/shell/SmartArtDialog").then((m) => ({ default: m.SmartArtDialog })))
const IconsDialog = lazy(() => import("@/shell/IconsDialog").then((m) => ({ default: m.IconsDialog })))
import { ThesaurusPane } from "@/shell/ThesaurusPane"
import { QueriesPane } from "@/shell/QueriesPane"
import { WorkbookLinksPane } from "@/shell/WorkbookLinksPane"
import { QueryPropertiesDialog } from "@/shell/QueryPropertiesDialog"
import { CreateNamesDialog, ErrorCheckingDialog, EvaluateFormulaDialog, NameManagerDialog } from "@/shell/AuditDialogs"
import { WatchWindow } from "@/shell/WatchWindow"
import { GroupDialog, OutlineSettingsDialog, SubtotalDialog } from "@/shell/OutlineDialogs"
import { DataTableDialog, GoalSeekDialog } from "@/shell/WhatIfDialogs"
import { ScenarioManagerDialog } from "@/shell/ScenarioDialogs"
import { AdvancedFilterDialog } from "@/shell/AdvancedFilterDialog"
import { ConsolidateDialog } from "@/shell/ConsolidateDialog"
import { ForecastDialog } from "@/shell/ForecastDialog"
import { ContextMenu, menuIcons, type MenuItem } from "@/shell/ContextMenu"
import { FormulaBar } from "@/shell/FormulaBar"
import { DefineNameDialog } from "@/shell/DefineNameDialog"
import { GoToDialog } from "@/shell/GoToDialog"
import { InsertFunctionDialog } from "@/shell/InsertFunctionDialog"
import type { Target } from "@/sheet/names"
import { InsertDeleteDialog } from "@/shell/InsertDeleteDialog"
import { SizeDialog } from "@/shell/SizeDialog"
import {
  isQuickKind,
  ManageRulesDialog,
  NewRuleDialog,
  QuickRuleDialog,
  type QuickKind,
} from "@/shell/cf/ConditionalDialogs"
import { MoveCopyDialog, UnhideSheetDialog } from "@/shell/SheetDialogs"
import {
  ProtectSheetDialog,
  ProtectWorkbookDialog,
  UnprotectSheetDialog,
  UnprotectWorkbookDialog,
} from "@/shell/ProtectDialogs"
import { SheetTabs } from "@/shell/SheetTabs"
import { StatusBar } from "@/shell/StatusBar"
import { ZoomDialog } from "@/shell/ZoomDialog"
import { QuickAccess, TellMe } from "@/shell/QuickAccess"
import { loadQat, saveQat } from "@/ribbon/qat-commands"
import { TitleBar } from "@/shell/TitleBar"
import { FillAcrossDialog } from "@/shell/FillAcrossDialog"
import { CustomViewsDialog } from "@/shell/CustomViewsDialog"
import { SignatureSetupDialog } from "@/shell/SignatureSetupDialog"
import { ObjectDialog } from "@/shell/ObjectDialog"
import { ArrangeWindowsDialog, CompareSideBySideDialog, UnhideWindowDialog } from "@/shell/WindowDialogs"
import { WorkbookStatisticsDialog } from "@/shell/WorkbookStatisticsDialog"
import { AltTextDialog } from "@/shell/AltTextDialog"
import { AllowEditRangesDialog, UnlockRangeDialog } from "@/shell/EditRangesDialogs"
import { loadDemo } from "./demo"
import { FileCommands, type SaveChoice } from "@/file/commands"
import { newDocument, type FileProblem, type WorkbookDocument } from "@/file/document"
import {
  FileNoticeDialog,
  FileProblemDialog,
  FileProgress,
  RecoveryDialog,
  SaveChangesDialog,
} from "@/file/FileDialogs"
import { folderOf, type RecentFile } from "@/file/recent"
import type { RecoveredFile } from "@/file/recovery"
import { TextImportWizard } from "@/file/TextImportWizard"
import { CreateTableDialog, ResizeTableDialog } from "@/shell/TableDialogs"
import { SelectDataDialog } from "@/shell/SelectDataDialog"
import type { CreateTableInit } from "@/sheet/tables"
import { LinkDialog } from "@/shell/LinkDialog"
import { SymbolDialog } from "@/shell/SymbolDialog"
import { FILE_LINK_WARNING, type LinkTarget } from "@/sheet/hyperlinks"
import type { TextImport } from "@cells/io"

const NOT_BUILT = "Not built yet."
const TABS_IN_TABLE = [...TABS, TABLE_DESIGN]
const TABS_IN_PIVOT = [...TABS, PIVOT_ANALYZE, PIVOT_DESIGN]
const TABS_IN_SPARKLINE = [...TABS, SPARKLINE]
const TABS_IN_SLICER = [...TABS, SLICER]
const TABS_IN_TIMELINE = [...TABS, TIMELINE]
const TABS_IN_TABLE_SPARKLINE = [...TABS, TABLE_DESIGN, SPARKLINE]
const TABS_IN_CHART = [...TABS, CHART_DESIGN]
const TABS_IN_PIVOT_CHART = [...TABS, PIVOT_CHART_ANALYZE, CHART_DESIGN]
const TABS_IN_SHAPE = [...TABS, SHAPE_FORMAT]
const TABS_IN_EQUATION = [...TABS, SHAPE_FORMAT, EQUATION]
const TABS_IN_PICTURE = [...TABS, PICTURE_FORMAT]
const TABS_IN_MODEL3D = [...TABS, MODEL3D_FORMAT]
const TABS_IN_SMARTART = [...TABS, SMARTART_DESIGN, SMARTART_FORMAT]

type DialogState =
  | { kind: "goto" }
  | { kind: "insert" | "delete" }
  | { kind: "size"; axis: "rows" | "cols"; indexes: number[]; initial: number | null }
  | { kind: "confirm"; message: string; ok: () => void }
  | { kind: "formatCells"; tab?: FormatTab }
  | {
      kind:
        | "moveCopySheet"
        | "unhideSheet"
        | "zoom"
        | "fillAcross"
        | "customViews"
        | "arrangeWindows"
        | "object"
        | "unhideWindow"
        | "compareSideBySide"
        | "statistics"
        | "spelling"
        | "icons"
        | "smartArt"
        | "allowEditRanges"
        | "cfManage"
        | "protectSheet"
        | "unprotectSheet"
        | "protectWorkbook"
        | "unprotectWorkbook"
        | "defineName"
        | "insertFunction"
        | "nameManager"
        | "createNames"
        | "errorChecking"
        | "evaluateFormula"
        | "outlineGroup"
        | "outlineUngroup"
        | "outlineSettings"
        | "subtotal"
        | "goalSeek"
        | "dataTable"
        | "scenarioManager"
        | "advancedFilter"
        | "consolidate"
        | "forecast"
        | "queryProperties"
        | "insertLink"
        | "symbol"
    }
  | { kind: "cfQuick"; quick: QuickKind }
  | { kind: "unlockRange"; name: string; then: () => void }
  | { kind: "sort"; range: Range }
  | { kind: "sortWarning"; then: (expand: boolean) => void; warning?: "sort" | "duplicates" }
  | { kind: "removeDuplicates"; range: Range }
  | { kind: "dataValidation"; init: ValidationDialogInit }
  | { kind: "createTable"; init: CreateTableInit }
  | { kind: "selectData" }
  | { kind: "altText"; chartId: number }
  | { kind: "signatureSetup"; editing?: number }
  | { kind: "pageSetup"; tab: PageSetupTab }
  | { kind: "resizeTable"; range: Range }
  | { kind: "sparklineData"; init: SparklineDialogInit }
  | { kind: "sparklineEmpty" }
  | { kind: "insertSlicers" }
  | { kind: "slicerSettings"; slicer: SheetSlicer }
  | { kind: "slicerConnections"; slicer?: SheetSlicer | SheetTimeline }
  | { kind: "insertTimelines" }
  | { kind: "sparklineDates" }
  | { kind: "sparklineAxis"; which: "min" | "max" }
  | { kind: "textToColumns"; range: Range; texts: string[] }
  | { kind: "filter"; request: FilterDialogRequest }
  | { kind: "pivotFilter"; request: PivotFilterRequest }
  | { kind: "recommendPivots"; source?: string }
  | { kind: "createPivot"; source?: string; mode?: "source" | "move" | "chart"; chartType?: string }
  | { kind: "pivotValue"; index: number }
  | { kind: "pivotShowAs"; index: number; as: PivotShowAs }
  | { kind: "pivotField"; field: number }
  | { kind: "pivotOptions" }
  | { kind: "pivotCalcField" }
  | { kind: "pivotGroup" }
  | { kind: "findReplace"; tab: "find" | "replace"; n: number }
  | { kind: "goToSpecial" }
  | { kind: "options"; page?: OptionsPage }
  | { kind: "ai"; tool: AiTool }
  | { kind: "cfNew"; start: string }
  | null

/** Light or dark as File › Options › Office Theme says: the system's setting, or always one of them. */
function useColorScheme() {
  const theme = useSettings((s) => s.theme)
  useEffect(() => {
    const query = window.matchMedia("(prefers-color-scheme: dark)")
    const apply = () => {
      document.documentElement.classList.toggle("dark", theme === "dark" || (theme === "system" && query.matches))
      // Android draws the status and navigation bars over a strip in the title bar's colour.
      syncSystemBars()
    }
    apply()
    query.addEventListener("change", apply)
    return () => query.removeEventListener("change", apply)
  }, [theme])
}

/**
 * Dates and numbers in the person's region (File › Options › Advanced): what typed text means and how cells show.
 * The sheet redraws when it changes.
 */
function useRegionalFormats(controller: SheetController) {
  const dateFormat = useSettings((s) => s.dateFormat)
  const decimalSeparator = useSettings((s) => s.decimalSeparator)
  const groupSeparator = useSettings((s) => s.groupSeparator)
  useEffect(() => {
    setRegional(regionalFrom({ dateFormat, decimalSeparator, groupSeparator }))
    controller.changed()
  }, [controller, dateFormat, decimalSeparator, groupSeparator])
}

/** Loads the settings shared with Inkspire, and again whenever the window comes back to the front. */
function useSharedSettings(port: SettingsPort) {
  useEffect(() => {
    void loadSettings(port)
    const again = () => void loadSettings(port)
    window.addEventListener("focus", again)
    return () => window.removeEventListener("focus", again)
  }, [port])
}

/**
 * The window: one workbook at a time, with File › New, Open, Save, Save As and Close, AutoRecover and Document
 * Recovery. Opening a file replaces the workbook (after asking to save changes); the workbook's own view below is
 * built anew for each one.
 */
export function App({ controller: given }: { controller?: SheetController } = {}) {
  const platform = usePlatform()
  const [first] = useState(() => newDocument(given))
  const [doc, setDoc] = useState<WorkbookDocument>(first)
  const [progress, setProgress] = useState<{ label: string; done: number | null } | null>(null)
  const [asking, setAsking] = useState<{ title: string; answer: (choice: SaveChoice) => void } | null>(null)
  const [problem, setProblem] = useState<FileProblem | null>(null)
  const [notice, setNotice] = useState<{ fileName: string; lines: string[] } | null>(null)
  const [recovered, setRecovered] = useState<RecoveredFile[]>([])
  const [recent, setRecent] = useState<RecentFile[]>([])
  const [importing, setImporting] = useState<{
    bytes: Uint8Array
    fileName: string
    answer: (choices: TextImport | null) => void
  } | null>(null)
  const [question, setQuestion] = useState<{ message: string; answer: (ok: boolean) => void } | null>(null)
  const [files] = useState(
    () =>
      new FileCommands(
        platform.files,
        {
          show: setDoc,
          askSave: (title) => new Promise((answer) => setAsking({ title, answer })),
          problem: setProblem,
          progress: setProgress,
          notice: (lines, fileName) => setNotice({ lines, fileName }),
          recentChanged: setRecent,
          textImport: (bytes, fileName) => new Promise((answer) => setImporting({ bytes, fileName, answer })),
          confirm: (message) => new Promise((answer) => setQuestion({ message, answer })),
        },
        first,
      ),
  )

  // cells:// links and files passed by the system: open the file, then go to the range.
  useEffect(
    () =>
      platform.opens.onRequest((items) => {
        void (async () => {
          for (const item of items) {
            const request = parseOpenRequest(item)
            if (!request) continue
            if (request.editFor) {
              // Saving back as it changes is only for the copies Inkspire makes; any other file opens as usual.
              if (await platform.links.isEditCopy(request.path)) await files.openEmbedded(request.path, request.editFor)
              else await files.openPath(request.path)
              continue
            }
            if (!(await files.openPath(request.path)) || !request.range) continue
            const c = files.doc.controller
            const target = c.names.resolve(request.range)
            if (target) c.names.goTo(target)
            else c.shell.alert(`The link points to ${request.range}, which isn't in ${files.doc.fileName}.`)
          }
        })()
      }),
    [platform, files],
  )

  // Inkspire's linked tables: what their ranges hold, written a moment after they change (docs/linked-ranges.md).
  const [live] = useState(
    () =>
      new LiveLinks(platform.links, () => ({
        path: files.doc.ref?.path ?? null,
        dirty: files.doc.dirty,
        controller: files.doc.controller,
      })),
  )
  useEffect(() => {
    live.schedule()
    // Edits, recalculation and saves (a first save gives the workbook its path) all come through here.
    const stopChanges = doc.controller.subscribe(() => live.schedule())
    // Inkspire may have linked something new while it had the focus.
    const onFocus = () => live.schedule()
    window.addEventListener("focus", onFocus)
    return () => {
      stopChanges()
      window.removeEventListener("focus", onFocus)
    }
  }, [live, doc])
  useEffect(() => () => live.stop(), [live])

  useEffect(() => {
    setRecent(files.recent)
    loadDemo(first.controller, new URLSearchParams(window.location.search).get("demo"))
    files.startAutoRecover()
    void files.recovered().then((list) => {
      if (list.length) setRecovered(list)
    })
    return () => files.stop()
  }, [files, first])

  // Closing the window asks about unsaved changes (a browser tab can only show its own prompt).
  useEffect(
    () =>
      platform.window.onCloseRequested(() => {
        if (!platform.window.controllable) return !files.doc.dirty
        return files.confirmClose().then(async (ok) => {
          if (ok) await files.discardRecovered({ id: files.doc.id } as RecoveredFile)
          return ok
        })
      }),
    [platform, files],
  )

  // A file dropped on the window opens, as in Excel (not while a dialog is open).
  useEffect(
    () =>
      platform.files.onDrop(({ ref, bytes }) => {
        if (document.querySelector('[aria-modal="true"]')) return
        void files.openBytes(bytes, ref)
      }),
    [platform, files],
  )

  // Excel's file shortcuts, from anywhere but an open dialog.
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (document.querySelector('[aria-modal="true"]')) return
      const primary = event.ctrlKey || event.metaKey
      const key = event.key.toLowerCase()
      let run: (() => Promise<boolean>) | null = null
      if (primary && !event.altKey && !event.shiftKey && key === "n") run = () => files.newWorkbook()
      else if ((primary && !event.altKey && key === "o") || (primary && key === "f12")) run = () => files.open()
      else if (primary && !event.altKey && !event.shiftKey && key === "s") run = () => files.save()
      else if (!primary && !event.altKey && key === "f12") run = () => files.saveAs()
      else if (primary && !event.altKey && (key === "w" || key === "f4")) run = () => files.close()
      if (!run) return
      event.preventDefault()
      event.stopPropagation()
      void run()
    }
    window.addEventListener("keydown", onKey, true)
    return () => window.removeEventListener("keydown", onKey, true)
  }, [files])

  return (
    <>
      <WorkbookWindow key={doc.id} doc={doc} files={files} recentFiles={recent} />
      {progress ? <FileProgress label={progress.label} done={progress.done} /> : null}
      {asking ? (
        <SaveChangesDialog
          title={asking.title}
          onChoose={(choice) => {
            setAsking(null)
            asking.answer(choice)
          }}
        />
      ) : null}
      {importing ? (
        <TextImportWizard
          bytes={importing.bytes}
          fileName={importing.fileName}
          onFinish={(choices) => {
            setImporting(null)
            importing.answer(choices)
          }}
          onCancel={() => {
            setImporting(null)
            importing.answer(null)
          }}
        />
      ) : null}
      {question ? (
        <Dialog
          title="Cells"
          width={440}
          onAccept={() => {
            setQuestion(null)
            question.answer(true)
          }}
          onCancel={() => {
            setQuestion(null)
            question.answer(false)
          }}
        >
          <p className="dialog-message">{question.message}</p>
        </Dialog>
      ) : null}
      {problem ? <FileProblemDialog problem={problem} onClose={() => setProblem(null)} /> : null}
      {notice && !problem ? (
        <FileNoticeDialog fileName={notice.fileName} lines={notice.lines} onClose={() => setNotice(null)} />
      ) : null}
      {recovered.length && !asking ? (
        <RecoveryDialog
          files={recovered}
          onOpen={(file) => {
            setRecovered((list) => list.filter((f) => f !== file))
            void files.openRecovered(file)
          }}
          onDiscard={(file) => {
            setRecovered((list) => list.filter((f) => f !== file))
            void files.discardRecovered(file)
          }}
          onClose={() => setRecovered([])}
        />
      ) : null}
    </>
  )
}

function WorkbookWindow({
  doc,
  files,
  recentFiles,
}: {
  doc: WorkbookDocument
  files: FileCommands
  recentFiles: RecentFile[]
}) {
  // View › New Window: the ribbon, formula bar, tabs and status bar follow the active window.
  const windowsVersion = useSyncExternalStore(doc.windows.subscribe, doc.windows.getVersion)
  const controller = doc.controller
  // &F and &Z in the header and footer Page Layout shows.
  const printFile = { fileName: doc.fileName, folder: folderOf(doc.ref?.path) }
  for (const w of doc.windows.all) w.printFile = printFile
  const platform = usePlatform()
  const [dialog, setDialog] = useState<DialogState>(null)
  const openDialogRef = useRef<(name: string) => void>(() => undefined)
  const settingsLoaded = useSettings((s) => s.loaded)
  const setupDone = useSettings((s) => s.setupDone)
  const [setupAgain, setSetupAgain] = useState(false)
  const [badEntry, setBadEntry] = useState<{ alert: ValidationAlert; answer: (a: AlertAnswer) => void } | null>(null)
  const [menu, setMenu] = useState<GridContextMenu | null>(null)
  // Excel's message boxes sit over whatever is open (a dialog stays open under them).
  const [alert, setAlert] = useState<{ message: string; then?: () => void; back: HTMLElement | null } | null>(null)
  const [recent, setRecent] = useState<string[]>([])
  // What's New (File › Account) opens by itself once after an update.
  const [backstage, setBackstage] = useState<string | null>(() => (whatsNewOnStart() ? "account" : null))
  const registry = useMemo(() => buildRegistry(), [])
  const [qat, setQatState] = useState(loadQat)
  const setQat = (ids: string[]) => {
    setQatState(ids)
    saveQat(ids)
  }
  // Right-click on a ribbon command: Excel's Add to Quick Access Toolbar.
  const [ribbonMenuAt, setRibbonMenuAt] = useState<{ x: number; y: number; id: string | null } | null>(null)
  const ribbonMenu = (event: React.MouseEvent) => {
    const target = event.target as HTMLElement
    if (target.closest("input, textarea, select")) return
    event.preventDefault()
    const control = target.closest<HTMLElement>(".ribbon-groups [data-command]")
    setRibbonMenuAt({ x: event.clientX, y: event.clientY, id: control?.dataset.command ?? null })
  }
  const ribbonHandle = useRef<RibbonHandle | null>(null)
  const shell = useShellKind()
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  useColorScheme()
  useSharedSettings(platform.settings)
  useRegionalFormats(controller)

  useEffect(() => {
    if (!platform.window.controllable) document.documentElement.classList.add("is-browser")
  }, [platform])

  // Android's Back closes what is open, else sends Cells to the background (android-files.ts).
  useEffect(() => installBackHandler(() => Boolean(controller.edit)), [controller])

  // Dialogs, menus and the File view take the whole screen on a phone (phone.css).
  useEffect(() => {
    document.documentElement.classList.toggle("is-phone", shell === "phone")
    controller.pivots.phone = shell === "phone"
  }, [shell, controller])

  useEffect(() => {
    const shell: SheetController["shell"] = {
      documentPath: () => files.doc.ref?.path ?? null,
      goTo: () => setDialog({ kind: "goto" }),
      insertDelete: ({ kind }) => setDialog({ kind }),
      confirm: (message, ok) => setDialog({ kind: "confirm", message, ok }),
      alert: (message, then) =>
        setAlert({
          message,
          then,
          back: document.activeElement instanceof HTMLElement ? document.activeElement : null,
        }),
      formatCells: (tab) => setDialog({ kind: "formatCells", tab }),
      insertFunction: () => setDialog({ kind: "insertFunction" }),
      openDialog: (name) => openDialogRef.current(name),
      sortWarning: (then, warning) => setDialog({ kind: "sortWarning", then, warning }),
      unlockRange: (name, then) => setDialog({ kind: "unlockRange", name, then }),
      removeDuplicates: (range) => setDialog({ kind: "removeDuplicates", range }),
      textToColumns: (range, texts) => setDialog({ kind: "textToColumns", range, texts }),
      sortDialog: (range) => setDialog({ kind: "sort", range }),
      filterDialog: (request) => setDialog({ kind: "filter", request }),
      pivotFilterDialog: (request) => setDialog({ kind: "pivotFilter", request }),
      findReplace: (tab) => setDialog({ kind: "findReplace", tab, n: Date.now() }),
      goToSpecial: () => setDialog({ kind: "goToSpecial" }),
      dataValidation: (init) => setDialog({ kind: "dataValidation", init }),
      validationAlert: (alert, answer) => setBadEntry({ alert, answer }),
      createTable: (init) => setDialog({ kind: "createTable", init }),
      chartData: () => setDialog({ kind: "selectData" }),
      showTab: (id) => ribbonHandle.current?.showTab(id),
      resizeTable: (range) => setDialog({ kind: "resizeTable", range }),
      sparklineDialog: (init) => setDialog({ kind: "sparklineData", init }),
      followLink: (target) => void followLink(target),
      openCellList: (pick) => window.dispatchEvent(new CustomEvent("cells:open-cell-list", { detail: { pick } })),
    }
    for (const w of doc.windows.all) w.shell = shell
  }, [controller, doc, windowsVersion])

  /** Double-click on an embedded file (Insert › Object): it opens in the app the system has for it. */
  const openObject = async (id: number) => {
    const o = controller.objects.find(id)
    const embed = o && !isChart(o) && o.kind === "picture" ? o.embed : undefined
    if (!embed) return
    try {
      const bytes = embed.data ?? embed.raw?.bytes
      if (!bytes || !platform.opens.openCopy) throw new Error("no way to open")
      await platform.opens.openCopy(embed.data ? embed.name : `${embed.name}.${embed.raw!.ext}`, bytes)
    } catch {
      controller.shell.alert(`Cells can't open ${embed.name}. There's no app on this device that opens it.`)
    }
  }

  /** A link out of the sheet: another workbook opens here (then goes to its place); the rest goes to the system. */
  const followLink = async (target: LinkTarget) => {
    if (target.kind === "workbook") {
      if (!(await files.openPath(target.path))) return
      if (target.location) files.doc.controller.hyperlinks.goTo(target.location)
      return
    }
    if (target.kind !== "external") return
    const open = () =>
      void platform.opens.external(target.target).catch(() => controller.shell.alert("Cannot open the specified file."))
    if (target.file) controller.shell.confirm(FILE_LINK_WARNING, open)
    else open()
  }

  const focusGrid = useCallback(() => {
    requestAnimationFrame(() => sheetEditor()?.focus({ preventScroll: true }))
  }, [])

  const closeDialog = useCallback(() => {
    setDialog(null)
    focusGrid()
  }, [focusGrid])

  /** A filter dialog works on the filter whose dropdown opened it (data.focus) until it closes. */
  const closeFilterDialog = useCallback(() => {
    controller.data.focus = null
    closeDialog()
  }, [controller, closeDialog])

  const copyToClipboard = async (cut: boolean) => {
    const content = controller.copy(cut)
    if (content) await platform.clipboard.write(content).catch(() => undefined)
  }

  const pasteFromClipboard = async () => {
    const content = await platform.clipboard.read()
    if (content) controller.paste(content.text, content.html)
  }

  const sizeDialog = (axis: "rows" | "cols") => {
    const indexes = controller.selectedIndexes(axis)
    const line = axis === "rows" ? controller.sheet.rows : controller.sheet.cols
    const sizes = new Set(indexes.map((i) => line.size(i)))
    const only = sizes.size === 1 ? [...sizes][0] : null
    setDialog({
      kind: "size",
      axis,
      indexes,
      initial: only === null ? null : axis === "rows" ? rowPoints(only) : columnChars(only),
    })
  }

  const menuItems = (target: GridContextMenu["target"]): MenuItem[] => {
    const key = (k: string) => (platform.os === "macos" ? `⌘${k}` : `Ctrl+${k}`)
    const items: MenuItem[] = [
      {
        kind: "item",
        label: "Cut",
        icon: menuIcons.cut,
        shortcut: key("X"),
        onSelect: () => void copyToClipboard(true),
      },
      {
        kind: "item",
        label: "Copy",
        icon: menuIcons.copy,
        shortcut: key("C"),
        onSelect: () => void copyToClipboard(false),
      },
      {
        kind: "item",
        label: "Paste",
        icon: menuIcons.paste,
        shortcut: key("V"),
        onSelect: () => void pasteFromClipboard(),
      },
      {
        kind: "item",
        label: "Paste Special...",
        disabled: !registry.isEnabled("PasteSpecialDialog", env),
        onSelect: () => registry.get("PasteSpecialDialog")?.run?.(env),
      },
      { kind: "separator" },
    ]
    // The cell menu's links: Link › Insert Link on a plain cell; Edit, Open and Remove on a linked one (Excel 365).
    const linkItems = (): MenuItem[] => {
      const { row, col } = controller.selection.active
      if (controller.hyperlinks.target(row, col)) {
        const own = controller.hyperlinks.activeLink() !== null
        return [
          { kind: "item", label: "Edit Hyperlink...", disabled: !own, onSelect: () => openDialog("insertLink") },
          { kind: "item", label: "Open Hyperlink", onSelect: () => void controller.hyperlinks.follow(row, col) },
          {
            kind: "item",
            label: "Remove Hyperlink",
            disabled: !own,
            onSelect: () => controller.hyperlinks.remove(true),
          },
        ]
      }
      return [
        {
          kind: "item",
          label: "Link",
          icon: menuIcons.link,
          items: [{ kind: "item", label: "Insert Link...", onSelect: () => openDialog("insertLink") }],
          onSelect: () => openDialog("insertLink"),
        },
      ]
    }
    const formatCells: MenuItem = {
      kind: "item",
      label: "Format Cells...",
      icon: menuIcons.formatCells,
      onSelect: () => controller.shell.formatCells(),
    }
    // Notes and comments: Excel's rows depend on what the cell has.
    const annotationItems = (): MenuItem[] => {
      const a = controller.annotations
      const { row, col } = controller.selection.active
      const merge = controller.sheet.merges.at(row, col)
      const at = merge ? { row: merge.r0, col: merge.c0 } : { row, col }
      const item = (label: string, onSelect: () => void): MenuItem => ({ kind: "item", label, onSelect })
      const has = a.at(at.row, at.col)
      if (has?.kind === "note")
        return [
          item("Edit Note", () => a.editNote(at.row, at.col)),
          item("Delete Note", () => a.deleteNotes()),
          item("Show/Hide Note", () => a.toggleNote()),
        ]
      if (has?.kind === "comment")
        return [
          item("New Reply", () => a.openThread(at.row, at.col)),
          item("Delete Comment", () => a.deleteThread(at.row, at.col)),
        ]
      return [item("New Comment", () => a.newComment()), item("New Note", () => a.newNote())]
    }
    const active = controller.selection.active
    const data = controller.data
    const act = (label: string, onSelect: () => void, disabled = false, icon?: ReactNode): MenuItem => ({
      kind: "item",
      label,
      disabled,
      icon,
      onSelect,
    })
    // The cell menu's Filter ›, by the active cell (Excel's order).
    const filterItems = (): MenuItem[] => {
      const header = data.filteredHeader()
      return [
        act(
          header === null ? "Clear Filter" : `Clear Filter From "${header}"`,
          () => void data.clearColumnFilter(),
          header === null,
        ),
        act("Reapply", () => void data.reapply(), data.autoFilter === null),
        { kind: "separator" },
        act("Filter by Selected Cell's Value", () => void data.filterBySelected("value")),
        act("Filter by Selected Cell's Color", () => void data.filterBySelected("cell")),
        act("Filter by Selected Cell's Font Color", () => void data.filterBySelected("font")),
        act("Filter by Selected Cell's Icon", () => void data.filterBySelected("icon"), data.activeIcon === null),
      ]
    }
    const sortItems = (): MenuItem[] => {
      const names = data.labels()
      return [
        act(names.ascending, () => data.quickSort(false), false, pictureFor("SortAscendingExcel", 16)),
        act(names.descending, () => data.quickSort(true), false, pictureFor("SortDescendingExcel", 16)),
        { kind: "separator" },
        act("Put Selected Cell Color On Top", () => void data.sortSelectedColorOnTop("cell")),
        act("Put Selected Font Color On Top", () => void data.sortSelectedColorOnTop("font")),
        act("Put Selected Cell Icon On Top", () => void data.sortSelectedColorOnTop("icon"), data.activeIcon === null),
        { kind: "separator" },
        act("Custom Sort...", () => data.openSortDialog()),
      ]
    }
    // A PivotTable's row or column label: Expand/Collapse (Excel's submenu).
    const pivotDrillItems = (): MenuItem[] => {
      const can = controller.pivots.drillState()
      if (!can) return []
      const drill = (scope: "item" | "field", collapse: boolean) => () => void controller.pivots.drill(scope, collapse)
      return [
        {
          kind: "item",
          label: "Expand/Collapse",
          onSelect: () => undefined,
          items: [
            act("Expand", drill("item", false), !can.expand, pictureFor("PivotExpandField", 16)),
            act("Collapse", drill("item", true), !can.collapse, pictureFor("PivotCollapseField", 16)),
            { kind: "separator" },
            act("Expand Entire Field", drill("field", false), !can.expandField, pictureFor("PivotExpandField", 16)),
            act(
              "Collapse Entire Field",
              drill("field", true),
              !can.collapseField,
              pictureFor("PivotCollapseField", 16),
            ),
          ],
        },
      ]
    }
    // A value cell's Summarize Values By and Show Values As (Excel's submenus, the current choice ticked).
    const pivotValueMenus = (index: number): MenuItem[] => {
      const v = controller.pivots.activeWithCache()?.table.values[index]
      if (!v) return []
      const tick = (on: boolean) => (on ? <Check size={15} strokeWidth={1.75} /> : undefined)
      const fns: PivotFn[] = ["sum", "count", "average", "max", "min", "product"]
      const summarize = (fn: PivotFn) => () =>
        void controller.pivots.setValueSettings(index, { name: v.name, fn, numFmt: v.numFmt })
      return [
        {
          kind: "item",
          label: "Summarize Values By",
          onSelect: () => undefined,
          items: [
            ...fns.map((fn) => act(PIVOT_FN_LABEL[fn], summarize(fn), false, tick(v.fn === fn))),
            { kind: "separator" },
            act("More Options...", () => openDialog(`pivotValue:${index}`)),
          ],
        },
        {
          kind: "item",
          label: "Show Values As",
          onSelect: () => undefined,
          items: SHOW_AS_CHOICES.map((c) => {
            const needsBase = c.as !== null && SHOW_AS_BASE_FIELD.includes(c.as)
            return act(
              needsBase ? `${c.label}...` : c.label,
              () =>
                needsBase
                  ? openDialog(`pivotShowAs:${index}/${c.as}`)
                  : void controller.pivots.setShowAs(index, { as: c.as }),
              false,
              tick((v.showAs ?? null) === c.as),
            )
          }),
        },
      ]
    }
    // Value Field Settings or Field Settings for the cell's field, then PivotTable Options, as Excel's cell menu has.
    function pivotSettingsItems(): MenuItem[] {
      const { row, col } = controller.selection.active
      const t = controller.pivots.settingsTargetAt(row, col)
      return [
        { kind: "separator" },
        ...(t?.kind === "value"
          ? [
              ...pivotValueMenus(t.index),
              { kind: "separator" as const },
              act("Value Field Settings...", () => openDialog(`pivotValue:${t.index}`)),
            ]
          : t?.kind === "field"
            ? [
                act("Group...", () => void controller.pivots.groupSelection(), !controller.pivots.groupTarget()),
                act("Ungroup...", () => void controller.pivots.ungroup(), !controller.pivots.canUngroup()),
                { kind: "separator" as const },
                act("Field Settings...", () => openDialog(`pivotField:${t.field}`)),
              ]
            : []),
        act("PivotTable Options...", () => openDialog("pivotOptions")),
        act(controller.pivots.fieldsOn() ? "Hide Field List" : "Show Field List", () =>
          controller.pivots.toggleFields(),
        ),
      ]
    }
    // Excel's lists (en-US), with what isn't built yet shown disabled.
    if (target === "cells") {
      items.push(
        { kind: "item", label: "Insert...", onSelect: () => controller.insertOrDelete("insert") },
        { kind: "item", label: "Delete...", onSelect: () => controller.insertOrDelete("delete") },
        { kind: "item", label: "Clear Contents", onSelect: () => controller.clearContents() },
        { kind: "separator" },
        ...(controller.pivots.active()
          ? ([
              {
                kind: "item",
                label: "Refresh",
                icon: pictureFor("RefreshAll", 16),
                onSelect: () => controller.pivots.refresh(),
              },
              ...pivotDrillItems(),
              ...(controller.pivots.canShowDetails(controller.selection.active.row, controller.selection.active.col)
                ? [
                    act("Show Details", () => {
                      const { row, col } = controller.selection.active
                      controller.pivots.showDetails(row, col)
                    }),
                  ]
                : []),
              ...pivotSettingsItems(),
              { kind: "separator" },
            ] as MenuItem[])
          : []),
        {
          kind: "item",
          label: "Quick Analysis",
          disabled: controller.quick.range() === null,
          // The menu hands focus back to the grid as it closes; the gallery opens after that.
          onSelect: () => void requestAnimationFrame(() => controller.quick.show()),
        },
        { kind: "separator" },
        { kind: "item", label: "Filter", items: filterItems(), onSelect: () => undefined },
        { kind: "item", label: "Sort", items: sortItems(), onSelect: () => undefined },
        { kind: "separator" },
        // Power Query, which Cells doesn't have (docs/next-todo.md, Phase 7).
        {
          kind: "item",
          label: "Get Data from Table/Range...",
          disabled: true,
          title: "Needs Power Query, which Cells doesn't have.",
          onSelect: () => undefined,
        },
        { kind: "separator" },
        ...annotationItems(),
        { kind: "separator" },
        formatCells,
        {
          kind: "item",
          label: "Pick From Drop-down List...",
          disabled: !(controller.cellList(active.row, active.col) ?? controller.pickList(active.row, active.col)),
          onSelect: () => void controller.openPickList(),
        },
        { kind: "item", label: "Define Name...", onSelect: () => openDialog("defineName") },
        { kind: "separator" },
        ...linkItems(),
      )
    } else {
      const axis = target
      items.push(
        { kind: "item", label: "Insert", onSelect: () => controller.insertLines(axis) },
        { kind: "item", label: "Delete", onSelect: () => controller.deleteLines(axis) },
        { kind: "item", label: "Clear Contents", onSelect: () => controller.clearContents() },
        { kind: "separator" },
        formatCells,
        {
          kind: "item",
          label: axis === "rows" ? "Row Height..." : "Column Width...",
          onSelect: () => sizeDialog(axis),
        },
        { kind: "item", label: "Hide", onSelect: () => controller.hideSelected(axis, true) },
        { kind: "item", label: "Unhide", onSelect: () => unhideAround(axis) },
      )
    }
    return items
  }

  /** Unhide acts on hidden lines inside the selection, as Excel does when you select across them. */
  const unhideAround = (axis: "rows" | "cols") => {
    const line = axis === "rows" ? controller.sheet.rows : controller.sheet.cols
    const hidden = controller.selectedIndexes(axis).filter((i) => line.isHidden(i))
    if (hidden.length) controller.setHidden(axis, hidden, false)
  }

  /** Zoom to Selection needs the size of the cell area on screen. */
  const zoomToSelection = () => {
    const host = document.querySelector<HTMLElement>(".grid-host")
    if (!host) return
    const header = controller.showHeadings ? 40 : 0
    controller.zoomToSelection(host.clientWidth - header - 14, host.clientHeight - header / 2 - 14)
  }

  const closeAlert = () => {
    const current = alert
    setAlert(null)
    if (current?.then) current.then()
    else if (dialog && current?.back?.isConnected) current.back.focus()
    else if (!dialog) focusGrid()
  }

  /** The ribbon's commands open the shell's dialogs by name (the controller's keys too, through the ref). */
  const openDialog = (name: string) => {
    const [kind, arg] = name.split(":")
    if (kind === "goto") setDialog({ kind: "goto" })
    else if (kind === "createPivot") setDialog({ kind: "createPivot", source: arg })
    else if (kind === "pivotSource") setDialog({ kind: "createPivot", mode: "source" })
    else if (kind === "pivotMove") setDialog({ kind: "createPivot", mode: "move" })
    else if (kind === "recommendPivots")
      setDialog({ kind: "recommendPivots", source: name.slice(kind.length + 1) || undefined })
    else if (kind === "createPivotChart") setDialog({ kind: "createPivot", mode: "chart", chartType: arg })
    else if (kind === "pivotValue") setDialog({ kind: "pivotValue", index: Number(arg) })
    else if (kind === "pivotShowAs") {
      const [index, as] = arg.split("/")
      setDialog({ kind: "pivotShowAs", index: Number(index), as: as as PivotShowAs })
    } else if (kind === "pivotField") setDialog({ kind: "pivotField", field: Number(arg) })
    else if (kind === "pivotOptions") setDialog({ kind: "pivotOptions" })
    else if (kind === "pivotCalcField") setDialog({ kind: "pivotCalcField" })
    else if (kind === "pivotGroup") setDialog({ kind: "pivotGroup" })
    else if (kind === "insert" || kind === "delete") setDialog({ kind })
    else if (kind === "size") sizeDialog(arg as "rows" | "cols")
    else if (kind === "autofit") controller.autofit?.(arg as "rows" | "cols")
    else if (kind === "toggle" && arg === "formulaBar") {
      controller.showFormulaBar = !controller.showFormulaBar
      controller.changed()
    } else if (kind === "formatCells") controller.shell.formatCells(arg as never)
    else if (
      kind === "moveCopySheet" ||
      kind === "unhideSheet" ||
      kind === "zoom" ||
      kind === "fillAcross" ||
      kind === "customViews" ||
      kind === "arrangeWindows" ||
      kind === "object" ||
      kind === "unhideWindow" ||
      kind === "compareSideBySide" ||
      kind === "statistics" ||
      kind === "spelling" ||
      kind === "icons" ||
      kind === "smartArt" ||
      kind === "allowEditRanges"
    )
      setDialog({ kind })
    else if (kind === "zoomToSelection") zoomToSelection()
    else if (kind === "sparklineEmpty" || kind === "sparklineDates") setDialog({ kind })
    else if (kind === "insertSlicers" || kind === "insertTimelines") setDialog({ kind })
    else if (kind === "timelineConnections") {
      const timeline = controller.timelines.active()
      if (timeline) setDialog({ kind: "slicerConnections", slicer: timeline })
    } else if (kind === "slicerSettings") {
      const slicer = controller.slicers.active()
      if (slicer) setDialog({ kind, slicer })
    } else if (kind === "slicerConnections") setDialog({ kind, slicer: controller.slicers.active() ?? undefined })
    else if (kind === "pivotFilterConnections") setDialog({ kind: "slicerConnections" })
    else if (kind === "sparklineAxis" && (arg === "min" || arg === "max")) setDialog({ kind, which: arg })
    else if (kind === "cf" && isQuickKind(arg)) setDialog({ kind: "cfQuick", quick: arg })
    else if (kind === "cfNew") setDialog({ kind: "cfNew", start: arg ?? "contains" })
    else if (kind === "cfManage") setDialog({ kind: "cfManage" })
    else if (kind === "ai" && (arg === "formula" || arg === "explain" || arg === "chart" || arg === "summary"))
      setDialog({ kind: "ai", tool: arg })
    else if (kind === "options") setDialog({ kind: "options", page: arg === "ai" ? "ai" : undefined })
    else if (kind === "altText" && arg) setDialog({ kind: "altText", chartId: Number(arg) })
    else if (kind === "openObject" && arg) void openObject(Number(arg))
    else if (kind === "signatureSetup") setDialog({ kind: "signatureSetup", editing: arg ? Number(arg) : undefined })
    else if (kind === "pageSetup") setDialog({ kind: "pageSetup", tab: (arg as PageSetupTab) ?? "page" })
    else if (
      kind === "protectSheet" ||
      kind === "unprotectSheet" ||
      kind === "protectWorkbook" ||
      kind === "unprotectWorkbook" ||
      kind === "defineName" ||
      kind === "insertFunction" ||
      kind === "nameManager" ||
      kind === "createNames" ||
      kind === "errorChecking" ||
      kind === "evaluateFormula" ||
      kind === "outlineGroup" ||
      kind === "outlineUngroup" ||
      kind === "outlineSettings" ||
      kind === "subtotal" ||
      kind === "goalSeek" ||
      kind === "dataTable" ||
      kind === "scenarioManager" ||
      kind === "advancedFilter" ||
      kind === "consolidate" ||
      kind === "forecast" ||
      kind === "queryProperties" ||
      kind === "insertLink" ||
      kind === "symbol"
    )
      setDialog({ kind })
  }

  openDialogRef.current = openDialog

  const env: RibbonEnv = {
    c: controller,
    copy: (cut) => void copyToClipboard(cut),
    paste: () => void pasteFromClipboard(),
    dialog: openDialog,
    toggleRibbon: () => ribbonHandle.current?.toggleCollapsed(),
    backstage: (page) => setBackstage(page ?? "home"),
    files,
  }

  // Ctrl+F1 shows or hides the ribbon's commands, from anywhere; Ctrl+P and Ctrl+F2 open File › Print.
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      const primary = event.ctrlKey || event.metaKey
      if (event.key === "F1" && primary) {
        event.preventDefault()
        ribbonHandle.current?.toggleCollapsed()
      } else if (
        event.key === "F5" &&
        event.altKey &&
        !event.shiftKey &&
        !document.querySelector('[aria-modal="true"]')
      ) {
        // Ctrl+Alt+F5: Refresh All; Alt+F5: Refresh the linked data at the active cell.
        event.preventDefault()
        event.stopPropagation()
        const c = files.doc.controller
        if (primary) void files.refreshAll().then(() => c.pivots.refreshAll())
        else if (c.pivots.active()) c.pivots.refresh()
        else {
          const q = c.data.activeQuery()
          if (q) void files.refreshQuery(q.sheet, q.name)
        }
      } else if (
        primary &&
        !event.altKey &&
        !event.shiftKey &&
        (event.key.toLowerCase() === "p" || event.key === "F2") &&
        !document.querySelector('[aria-modal="true"]')
      ) {
        event.preventDefault()
        event.stopPropagation()
        setBackstage("print")
      } else if (event.key === "F6" && primary && !event.altKey && !document.querySelector('[aria-modal="true"]')) {
        // Ctrl+F6 / Ctrl+Shift+F6: the workbook's next or previous window (View › New Window).
        if (files.doc.windows.visible.length < 2) return
        event.preventDefault()
        event.stopPropagation()
        files.doc.windows.cycle(event.shiftKey ? -1 : 1)
      }
    }
    window.addEventListener("keydown", onKey)
    return () => window.removeEventListener("keydown", onKey)
  }, [])

  // The macOS menu bar (src-tauri/src/menu.rs): each item runs the command it names, as the ribbon would.
  const onMenuCommand = useRef<(id: string) => void>(() => {})
  onMenuCommand.current = (id) => {
    const action = menuBarAction(id, isTextField(document.activeElement, Boolean(controller.edit)))
    if (!action) return
    if (action.kind === "command") {
      if (registry.isEnabled(action.id, env)) registry.get(action.id)?.run?.(env)
    } else if (action.kind === "text") document.execCommand(action.command)
    else if (action.kind === "selectAll") controller.selectAllOrRegion()
    else if (action.kind === "backstage") setBackstage(action.page)
    else if (action.kind === "dialog") openDialog(action.name)
    else if (action.kind === "saveAs") void files.saveAs()
    else if (action.kind === "close") void files.close()
    else if (action.kind === "toggleRibbon") ribbonHandle.current?.toggleCollapsed()
    else void platform.window.close()
  }
  useEffect(() => platform.window.onMenuCommand?.((id) => onMenuCommand.current(id)), [platform])

  const goTo = (target: Target, text: string) => {
    const current = rangeName(controller.selection.ranges[controller.selection.ranges.length - 1])
    setRecent((items) => [current, ...items.filter((item) => item !== current && item !== text)].slice(0, 10))
    controller.names.goTo(target)
    closeDialog()
  }

  const formulaBar = controller.showFormulaBar ? (
    <FormulaBar
      controller={controller}
      dragging={false}
      onFocusGrid={focusGrid}
      onInsertFunction={() => openDialog("insertFunction")}
    />
  ) : null
  const accessibilityPane = (
    <AccessibilityPane controller={controller} onDone={focusGrid} onDialog={(name) => openDialog(name)} />
  )
  const linksPane = <WorkbookLinksPane controller={controller} files={files} onDone={focusGrid} />
  const queriesPane = (
    <QueriesPane
      controller={controller}
      files={files}
      onProperties={() => openDialog("queryProperties")}
      onDone={focusGrid}
    />
  )
  const clipboardPane = (
    <ClipboardPane controller={controller} onPaste={(text, html) => controller.paste(text, html)} onDone={focusGrid} />
  )
  const grid = (
    <WindowsArea
      windows={doc.windows}
      baseTitle={doc.title}
      grid={(c, active) => <GridView key={doc.windows.key(c)} controller={c} active={active} onContextMenu={setMenu} />}
    />
  )
  const drawings = controller.objects.drawings()
  const smartArt = controller.objects.smartArt()
  const smartArtPane =
    smartArt && controller.objects.textPaneOpen ? (
      <SmartArtTextPane key={smartArt.id} controller={controller} onDone={focusGrid} />
    ) : null
  const ribbonTabs = controller.charts.activePivot()
    ? TABS_IN_PIVOT_CHART
    : controller.charts.active()
      ? TABS_IN_CHART
      : smartArt
        ? TABS_IN_SMARTART
        : drawings.length === 1 && drawings[0].kind === "picture" && drawings[0].model3d
          ? TABS_IN_MODEL3D
          : drawings.length && drawings.every((o) => o.kind === "slicer")
            ? TABS_IN_SLICER
            : drawings.length && drawings.every((o) => o.kind === "timeline")
              ? TABS_IN_TIMELINE
              : drawings.length && drawings.every((o) => o.kind === "picture")
                ? TABS_IN_PICTURE
                : drawings.length
                  ? drawings.length === 1 && isEquation(drawings[0])
                    ? TABS_IN_EQUATION
                    : TABS_IN_SHAPE
                  : controller.sparklines.active()
                    ? controller.tables.active()
                      ? TABS_IN_TABLE_SPARKLINE
                      : TABS_IN_SPARKLINE
                    : controller.tables.active()
                      ? TABS_IN_TABLE
                      : controller.pivots.active()
                        ? TABS_IN_PIVOT
                        : TABS
  // The phone's palette: the ribbon's tabs but Help (its commands are in the File view's About).
  const paletteTabs = ribbonTabs.filter((tab) => tab !== HELP)

  return (
    <TooltipProvider>
      <div className={shell === "phone" ? "app-window-frame is-phone" : "app-window-frame"}>
        {shell === "phone" ? (
          <PhoneShell
            controller={controller}
            env={env}
            registry={registry}
            tabs={paletteTabs}
            title={controller.group.active ? `${doc.windowTitle} [Group]` : doc.windowTitle}
            formulaBar={formulaBar}
            onFile={() => setBackstage("home")}
            onFocusGrid={focusGrid}
          >
            {grid}
          </PhoneShell>
        ) : (
          <>
            <TitleBar
              controller={controller}
              title={`${doc.windowTitle}${controller.group.active ? " [Group]" : ""} - Cells`}
              qat={<QuickAccess registry={registry} env={env} ids={qat} onChange={setQat} onDone={focusGrid} />}
              search={
                <TellMe
                  registry={registry}
                  env={env}
                  onShow={(id) => ribbonHandle.current?.show(id) ?? false}
                  onDone={focusGrid}
                />
              }
            />
            <div className="ribbon-host" onContextMenu={ribbonMenu}>
              <Ribbon
                env={env}
                registry={registry}
                tabs={ribbonTabs}
                onFile={() => setBackstage("home")}
                qatKeytips={qat.slice(0, 9).map((id, i) => ({
                  key: String(i + 1),
                  run: () => {
                    if (registry.isEnabled(id, env)) registry.get(id)?.run?.(env)
                  },
                }))}
                handle={(h) => (ribbonHandle.current = h)}
              />
            </div>
            {ribbonMenuAt ? (
              <ContextMenu
                x={ribbonMenuAt.x}
                y={ribbonMenuAt.y}
                onClose={() => setRibbonMenuAt(null)}
                items={[
                  {
                    kind: "item",
                    label: "Add to Quick Access Toolbar",
                    disabled: !ribbonMenuAt.id || qat.includes(ribbonMenuAt.id),
                    onSelect: () => ribbonMenuAt.id && setQat([...qat, ribbonMenuAt.id]),
                  },
                  { kind: "separator" },
                  {
                    kind: "item",
                    label: "Customize the Ribbon...",
                    disabled: true,
                    title: NOT_BUILT,
                    onSelect: () => undefined,
                  },
                  {
                    kind: "item",
                    label: "Collapse the Ribbon",
                    onSelect: () => ribbonHandle.current?.toggleCollapsed(),
                  },
                ]}
              />
            ) : null}
            {formulaBar}
            <div className="sheet-row">
              {controller.clipboardPaneOpen ? clipboardPane : null}
              <main className="sheet-area">{grid}</main>
              {controller.annotations.paneOpen ? <CommentsPane controller={controller} onDone={focusGrid} /> : null}
              {controller.accessibilityOpen ? accessibilityPane : null}
              {controller.thesaurusOpen ? <ThesaurusPane controller={controller} onDone={focusGrid} /> : null}
              {controller.objects.paneOpen ? <SelectionPane controller={controller} onDone={focusGrid} /> : null}
              {smartArtPane}
              {controller.audit.watchOpen ? <WatchWindow controller={controller} onDone={focusGrid} /> : null}
              {controller.data.queriesOpen ? queriesPane : null}
              {controller.data.linksOpen ? linksPane : null}
              {controller.pivots.fieldsShown(false) ? (
                <PivotFieldsPane controller={controller} onDone={focusGrid} onDialog={openDialog} />
              ) : null}
            </div>
            <SheetTabs controller={controller} onDialog={openDialog} onFocusGrid={focusGrid} />
            <StatusBar controller={controller} onZoomDialog={() => openDialog("zoom")} />
          </>
        )}
        {shell === "phone" && controller.clipboardPaneOpen ? clipboardPane : null}
        {shell === "phone" && controller.data.queriesOpen ? queriesPane : null}
        {shell === "phone" && controller.data.linksOpen ? linksPane : null}
        {shell === "phone" && controller.audit.watchOpen ? (
          <WatchWindow controller={controller} onDone={focusGrid} />
        ) : null}
        {shell === "phone" && controller.accessibilityOpen ? accessibilityPane : null}
        {shell === "phone" ? smartArtPane : null}
        {shell === "phone" && controller.objects.paneOpen ? (
          <SelectionPane controller={controller} onDone={focusGrid} />
        ) : null}
        {shell === "phone" && controller.thesaurusOpen ? (
          <ThesaurusPane controller={controller} onDone={focusGrid} />
        ) : null}
        {shell === "phone" && controller.pivots.fieldsShown(true) ? (
          <PivotFieldsPane controller={controller} onDone={focusGrid} onDialog={openDialog} />
        ) : null}
        {shell === "phone" && controller.annotations.paneOpen ? (
          <CommentsPane controller={controller} onDone={focusGrid} />
        ) : null}
        {menu && shell === "phone" ? (
          <ActionSheet
            label={{ cells: "Cells", rows: "Rows", cols: "Columns" }[menu.target]}
            items={menuItems(menu.target)}
            onClose={() => (setMenu(null), focusGrid())}
          />
        ) : menu ? (
          <ContextMenu
            x={menu.x}
            y={menu.y}
            items={menuItems(menu.target)}
            onClose={() => (setMenu(null), focusGrid())}
          />
        ) : null}
        {dialog?.kind === "goto" ? (
          <GoToDialog
            names={controller.names.all().map((n) => n.name)}
            recent={recent}
            resolve={(text) => controller.names.resolve(text)}
            onGo={goTo}
            onSpecial={() => setDialog({ kind: "goToSpecial" })}
            onClose={closeDialog}
          />
        ) : null}
        {dialog?.kind === "insert" || dialog?.kind === "delete" ? (
          <InsertDeleteDialog
            kind={dialog.kind}
            onClose={closeDialog}
            onChoose={(choice) => {
              const kind = dialog.kind as "insert" | "delete"
              if (choice === "row") kind === "insert" ? controller.insertLines("rows") : controller.deleteLines("rows")
              else if (choice === "col")
                kind === "insert" ? controller.insertLines("cols") : controller.deleteLines("cols")
              else controller.shiftCells(kind, choice)
              closeDialog()
            }}
          />
        ) : null}
        {dialog?.kind === "formatCells" ? (
          <FormatCellsDialog controller={controller} tab={dialog.tab} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "confirm" ? (
          <Dialog
            title="Cells"
            width={360}
            onCancel={closeDialog}
            onAccept={() => {
              closeDialog()
              dialog.ok()
            }}
          >
            <p className="dialog-message">{dialog.message}</p>
          </Dialog>
        ) : null}
        {dialog?.kind === "moveCopySheet" ? <MoveCopyDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "unhideSheet" ? <UnhideSheetDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "cfQuick" ? (
          <QuickRuleDialog controller={controller} kind={dialog.quick} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "cfNew" ? (
          <NewRuleDialog controller={controller} start={dialog.start} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "defineName" ? <DefineNameDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "insertLink" ? <LinkDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "symbol" ? <SymbolDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "nameManager" ? <NameManagerDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "createNames" ? <CreateNamesDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "outlineGroup" || dialog?.kind === "outlineUngroup" ? (
          <GroupDialog controller={controller} ungroup={dialog.kind === "outlineUngroup"} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "outlineSettings" ? (
          <OutlineSettingsDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "subtotal" ? <SubtotalDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "goalSeek" ? <GoalSeekDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "dataTable" ? <DataTableDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "advancedFilter" ? (
          <AdvancedFilterDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "consolidate" ? <ConsolidateDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "queryProperties" ? (
          <QueryPropertiesDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "forecast" ? <ForecastDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "scenarioManager" ? (
          <ScenarioManagerDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "evaluateFormula" ? (
          <EvaluateFormulaDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "errorChecking" ? (
          <ErrorCheckingDialog
            controller={controller}
            onClose={closeDialog}
            onEvaluate={() => setDialog({ kind: "evaluateFormula" })}
          />
        ) : null}
        {dialog?.kind === "sort" ? (
          <SortDialog controller={controller} range={dialog.range} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "findReplace" ? (
          <FindReplaceDialog key={dialog.n} controller={controller} tab={dialog.tab} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "goToSpecial" ? <GoToSpecialDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "filter" && dialog.request.kind === "custom" ? (
          <CustomFilterDialog
            controller={controller}
            col={dialog.request.col}
            preset={dialog.request.preset}
            onClose={closeFilterDialog}
          />
        ) : null}
        {dialog?.kind === "filter" && dialog.request.kind === "top10" ? (
          <Top10Dialog controller={controller} col={dialog.request.col} onClose={closeFilterDialog} />
        ) : null}
        {dialog?.kind === "pivotValue" ? (
          <ValueFieldSettingsDialog controller={controller} index={dialog.index} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "pivotShowAs" ? (
          <ShowValuesAsDialog controller={controller} index={dialog.index} as={dialog.as} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "pivotField" ? (
          <FieldSettingsDialog controller={controller} field={dialog.field} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "pivotOptions" ? <PivotOptionsDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "pivotGroup" ? <GroupingDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "pivotCalcField" ? (
          <CalculatedFieldDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "createPivot" ? (
          <CreatePivotDialog
            controller={controller}
            source={dialog.source}
            mode={dialog.mode}
            chartType={dialog.chartType}
            onClose={closeDialog}
          />
        ) : null}
        {dialog?.kind === "recommendPivots" ? (
          <RecommendedPivotsDialog controller={controller} source={dialog.source} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "pivotFilter" ? (
          <PivotFilterDialog controller={controller} request={dialog.request} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "textToColumns" ? (
          <TextImportWizard
            cells={dialog.texts}
            destination={absoluteName({ row: dialog.range.r0, col: dialog.range.c0 })}
            initial={controller.data.textColumns}
            onCancel={closeDialog}
            onFinish={(choices, destination) => {
              const at = parseCell(destination.replace(/\$/g, "").trim())
              if (!at) {
                controller.shell.alert("The destination reference is not valid.")
                return
              }
              const range = dialog.range
              closeDialog()
              controller.data.textToColumns(range, choices, at)
            }}
          />
        ) : null}
        {dialog?.kind === "dataValidation" ? (
          <DataValidationDialog controller={controller} init={dialog.init} onClose={closeDialog} />
        ) : null}
        {badEntry ? (
          <ValidationAlertDialog
            alert={badEntry.alert}
            onAnswer={(answer) => {
              setBadEntry(null)
              badEntry.answer(answer)
              focusGrid()
            }}
          />
        ) : null}
        {dialog?.kind === "signatureSetup" ? (
          <SignatureSetupDialog controller={controller} editing={dialog.editing} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "altText" ? (
          <AltTextDialog controller={controller} chartId={dialog.chartId} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "selectData" ? <SelectDataDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "pageSetup" ? (
          <PageSetupDialog controller={controller} tab={dialog.tab} fileName={doc.fileName} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "createTable" ? (
          <CreateTableDialog controller={controller} init={dialog.init} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "sparklineData" ? (
          <SparklineDataDialog controller={controller} init={dialog.init} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "insertSlicers" ? (
          <InsertSlicersDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "insertTimelines" ? (
          <InsertTimelinesDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "slicerConnections" ? (
          <SlicerConnectionsDialog controller={controller} slicer={dialog.slicer} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "slicerSettings" ? (
          <SlicerSettingsDialog controller={controller} slicer={dialog.slicer} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "sparklineEmpty" ? (
          <SparklineEmptyDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "sparklineDates" ? (
          <SparklineDateDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "sparklineAxis" ? (
          <SparklineAxisDialog controller={controller} which={dialog.which} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "resizeTable" ? (
          <ResizeTableDialog controller={controller} range={dialog.range} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "removeDuplicates" ? (
          <RemoveDuplicatesDialog controller={controller} range={dialog.range} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "sortWarning" ? (
          <SortWarningDialog
            kind={dialog.warning}
            onClose={closeDialog}
            onChoose={(expand) => {
              setDialog(null)
              dialog.then(expand)
              focusGrid()
            }}
          />
        ) : null}
        {dialog?.kind === "insertFunction" ? (
          <InsertFunctionDialog
            onClose={closeDialog}
            onInsert={(name) => {
              setDialog(null)
              controller.insertFunction(name)
              focusGrid()
            }}
          />
        ) : null}
        {dialog?.kind === "protectSheet" ? <ProtectSheetDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "protectWorkbook" ? (
          <ProtectWorkbookDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "unprotectWorkbook" ? (
          <UnprotectWorkbookDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "unprotectSheet" ? (
          <UnprotectSheetDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "cfManage" ? <ManageRulesDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "ai" ? (
          <AiDialog
            tool={dialog.tool}
            controller={controller}
            onClose={closeDialog}
            onOptions={() => setDialog({ kind: "options", page: "ai" })}
          />
        ) : null}
        {dialog?.kind === "options" ? (
          <OptionsDialog
            page={dialog.page}
            onClose={closeDialog}
            onRunSetup={() => {
              closeDialog()
              setSetupAgain(true)
            }}
          />
        ) : null}
        {settingsLoaded && (!setupDone || setupAgain) && !backstage ? (
          <FirstRunSetup
            onDone={() => {
              setSetupAgain(false)
              focusGrid()
            }}
          />
        ) : null}
        {dialog?.kind === "allowEditRanges" ? (
          <AllowEditRangesDialog
            controller={controller}
            onClose={closeDialog}
            onProtect={() => setDialog({ kind: "protectSheet" })}
          />
        ) : null}
        {dialog?.kind === "unlockRange" ? (
          <UnlockRangeDialog
            controller={controller}
            name={dialog.name}
            onUnlocked={dialog.then}
            onClose={closeDialog}
          />
        ) : null}
        {dialog?.kind === "spelling" ? <SpellingDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "smartArt" ? (
          <Suspense fallback={null}>
            <SmartArtDialog controller={controller} onClose={closeDialog} onInserted={() => setDialog(null)} />
          </Suspense>
        ) : null}
        {dialog?.kind === "icons" ? (
          <Suspense fallback={null}>
            <IconsDialog controller={controller} onClose={closeDialog} />
          </Suspense>
        ) : null}
        {dialog?.kind === "statistics" ? (
          <WorkbookStatisticsDialog controller={controller} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "customViews" ? <CustomViewsDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "object" ? <ObjectDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "arrangeWindows" ? (
          <ArrangeWindowsDialog windows={doc.windows} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "unhideWindow" ? (
          <UnhideWindowDialog windows={doc.windows} baseTitle={doc.title} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "compareSideBySide" ? (
          <CompareSideBySideDialog windows={doc.windows} baseTitle={doc.title} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "fillAcross" ? (
          <FillAcrossDialog onFill={(what) => controller.group.fillAcross(what)} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "zoom" ? (
          <ZoomDialog
            zoom={controller.zoom}
            onZoom={(z) => controller.setZoom(z)}
            onFitSelection={zoomToSelection}
            onClose={closeDialog}
          />
        ) : null}
        {dialog?.kind === "size" ? (
          <SizeDialog
            axis={dialog.axis}
            initial={dialog.initial}
            onClose={closeDialog}
            onApply={(value) => {
              const px = dialog.axis === "rows" ? pointsToPixels(value) : charsToPixels(value)
              controller.setLineSize(dialog.axis, dialog.indexes, px)
              closeDialog()
            }}
          />
        ) : null}
        {backstage ? (
          <Backstage
            page={backstage}
            files={files}
            controller={controller}
            onPageSetup={(tab) => setDialog({ kind: "pageSetup", tab })}
            onOptions={() => {
              setBackstage(null)
              setDialog({ kind: "options" })
            }}
            recent={recentFiles}
            file={{
              title: doc.title,
              fileName: doc.fileName,
              folder: folderOf(doc.ref?.path),
              saved: !doc.dirty,
              hasFile: Boolean(doc.ref?.path || doc.ref?.handle),
            }}
            onClose={() => {
              setBackstage(null)
              focusGrid()
            }}
          />
        ) : null}
        {alert ? (
          <Dialog
            title="Cells"
            width={400}
            cancelLabel={null}
            onCancel={() => closeAlert()}
            onAccept={() => closeAlert()}
          >
            <p className="dialog-message">{alert.message}</p>
          </Dialog>
        ) : null}
      </div>
    </TooltipProvider>
  )
}
