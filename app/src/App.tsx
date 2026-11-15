import { PageSetupDialog, type PageSetupTab } from "@/shell/PageSetupDialog"
import { parseCell, rangeName, type Range } from "@cells/core"
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react"
import { columnChars, charsToPixels, pointsToPixels, rowPoints } from "@/grid/autofit"
import { GridView, type GridContextMenu } from "@/grid/GridView"
import { usePlatform } from "@/platform/context"
import type { SettingsPort } from "@cells/platform"
import { loadSettings, useSettings } from "@/settings/settings"
import { OptionsDialog } from "@/settings/OptionsDialog"
import { FirstRunSetup } from "@/settings/FirstRunSetup"
import { SheetController } from "@/sheet/controller"
import { absoluteName } from "@/sheet/text-columns"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { Ribbon, type RibbonHandle } from "@/ribbon/Ribbon"
import { CHART_DESIGN, TABLE_DESIGN, TABS } from "@/ribbon/tabs"
import { Backstage } from "@/shell/Backstage"
import { whatsNewOnStart } from "@/about/release-notes"
import { Dialog } from "@/shell/Dialog"
import { RemoveDuplicatesDialog } from "@/shell/RemoveDuplicatesDialog"
import { DataValidationDialog, ValidationAlertDialog } from "@/shell/DataValidationDialog"
import type { AlertAnswer, ValidationAlert, ValidationDialogInit } from "@/sheet/validation"
import { FindReplaceDialog, GoToSpecialDialog } from "@/shell/FindDialogs"
import { CustomFilterDialog, Top10Dialog } from "@/shell/FilterDialogs"
import type { FilterDialogRequest } from "@/sheet/data"
import { SortDialog, SortWarningDialog } from "@/shell/SortDialog"
import { FormatCellsDialog, type FormatTab } from "@/shell/format-cells/FormatCellsDialog"
import { TooltipProvider } from "@/ui/tooltip"
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
import type { TextImport } from "@cells/io"

const NOT_BUILT = "Not built yet."
const TABS_IN_TABLE = [...TABS, TABLE_DESIGN]
const TABS_IN_CHART = [...TABS, CHART_DESIGN]

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
        | "cfManage"
        | "protectSheet"
        | "unprotectSheet"
        | "protectWorkbook"
        | "unprotectWorkbook"
        | "defineName"
        | "insertFunction"
    }
  | { kind: "cfQuick"; quick: QuickKind }
  | { kind: "sort"; range: Range }
  | { kind: "sortWarning"; then: (expand: boolean) => void; warning?: "sort" | "duplicates" }
  | { kind: "removeDuplicates"; range: Range }
  | { kind: "dataValidation"; init: ValidationDialogInit }
  | { kind: "createTable"; init: CreateTableInit }
  | { kind: "selectData" }
  | { kind: "pageSetup"; tab: PageSetupTab }
  | { kind: "resizeTable"; range: Range }
  | { kind: "textToColumns"; range: Range; texts: string[] }
  | { kind: "filter"; request: FilterDialogRequest }
  | { kind: "findReplace"; tab: "find" | "replace"; n: number }
  | { kind: "goToSpecial" }
  | { kind: "options" }
  | { kind: "cfNew"; start: string }
  | null

/** Light or dark as File › Options › Office Theme says: the system's setting, or always one of them. */
function useColorScheme() {
  const theme = useSettings((s) => s.theme)
  useEffect(() => {
    const query = window.matchMedia("(prefers-color-scheme: dark)")
    const apply = () =>
      document.documentElement.classList.toggle("dark", theme === "dark" || (theme === "system" && query.matches))
    apply()
    query.addEventListener("change", apply)
    return () => query.removeEventListener("change", apply)
  }, [theme])
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
  const controller = doc.controller
  const platform = usePlatform()
  const [dialog, setDialog] = useState<DialogState>(null)
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
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  useColorScheme()
  useSharedSettings(platform.settings)

  useEffect(() => {
    if (!platform.window.controllable) document.documentElement.classList.add("is-browser")
  }, [platform])

  useEffect(() => {
    controller.shell = {
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
      sortWarning: (then, warning) => setDialog({ kind: "sortWarning", then, warning }),
      removeDuplicates: (range) => setDialog({ kind: "removeDuplicates", range }),
      textToColumns: (range, texts) => setDialog({ kind: "textToColumns", range, texts }),
      sortDialog: (range) => setDialog({ kind: "sort", range }),
      filterDialog: (request) => setDialog({ kind: "filter", request }),
      findReplace: (tab) => setDialog({ kind: "findReplace", tab, n: Date.now() }),
      goToSpecial: () => setDialog({ kind: "goToSpecial" }),
      dataValidation: (init) => setDialog({ kind: "dataValidation", init }),
      validationAlert: (alert, answer) => setBadEntry({ alert, answer }),
      createTable: (init) => setDialog({ kind: "createTable", init }),
      chartData: () => setDialog({ kind: "selectData" }),
      showTab: (id) => ribbonHandle.current?.showTab(id),
      resizeTable: (range) => setDialog({ kind: "resizeTable", range }),
      openCellList: () => window.dispatchEvent(new CustomEvent("cells:open-cell-list")),
    }
  }, [controller])

  const focusGrid = useCallback(() => {
    requestAnimationFrame(() =>
      document.querySelector<HTMLTextAreaElement>(".cell-editor")?.focus({ preventScroll: true }),
    )
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
        disabled: true,
        title: NOT_BUILT,
        onSelect: () => undefined,
      },
      { kind: "separator" },
    ]
    const later = (label: string, more = false): MenuItem => ({
      kind: "item",
      label,
      more,
      disabled: true,
      title: NOT_BUILT,
      onSelect: () => undefined,
    })
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
    // Excel's lists (en-US), with what isn't built yet shown disabled.
    if (target === "cells") {
      items.push(
        { kind: "item", label: "Insert...", onSelect: () => controller.insertOrDelete("insert") },
        { kind: "item", label: "Delete...", onSelect: () => controller.insertOrDelete("delete") },
        { kind: "item", label: "Clear Contents", onSelect: () => controller.clearContents() },
        { kind: "separator" },
        later("Quick Analysis"),
        { kind: "separator" },
        later("Filter", true),
        later("Sort", true),
        { kind: "separator" },
        later("Get Data from Table/Range..."),
        { kind: "separator" },
        ...annotationItems(),
        { kind: "separator" },
        formatCells,
        later("Pick From Drop-down List..."),
        { kind: "item", label: "Define Name...", onSelect: () => openDialog("defineName") },
        { kind: "separator" },
        later("Link", true),
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

  /** The ribbon's commands open the shell's dialogs by name. */
  const openDialog = (name: string) => {
    const [kind, arg] = name.split(":")
    if (kind === "goto") setDialog({ kind: "goto" })
    else if (kind === "insert" || kind === "delete") setDialog({ kind })
    else if (kind === "size") sizeDialog(arg as "rows" | "cols")
    else if (kind === "autofit") controller.autofit?.(arg as "rows" | "cols")
    else if (kind === "toggle" && arg === "formulaBar") {
      controller.showFormulaBar = !controller.showFormulaBar
      controller.changed()
    } else if (kind === "formatCells") controller.shell.formatCells(arg as never)
    else if (kind === "moveCopySheet" || kind === "unhideSheet" || kind === "zoom") setDialog({ kind })
    else if (kind === "zoomToSelection") zoomToSelection()
    else if (kind === "cf" && isQuickKind(arg)) setDialog({ kind: "cfQuick", quick: arg })
    else if (kind === "cfNew") setDialog({ kind: "cfNew", start: arg ?? "contains" })
    else if (kind === "cfManage") setDialog({ kind: "cfManage" })
    else if (kind === "pageSetup") setDialog({ kind: "pageSetup", tab: (arg as PageSetupTab) ?? "page" })
    else if (
      kind === "protectSheet" ||
      kind === "unprotectSheet" ||
      kind === "protectWorkbook" ||
      kind === "unprotectWorkbook" ||
      kind === "defineName" ||
      kind === "insertFunction"
    )
      setDialog({ kind })
  }

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
        primary &&
        !event.altKey &&
        !event.shiftKey &&
        (event.key.toLowerCase() === "p" || event.key === "F2") &&
        !document.querySelector('[aria-modal="true"]')
      ) {
        event.preventDefault()
        event.stopPropagation()
        setBackstage("print")
      }
    }
    window.addEventListener("keydown", onKey)
    return () => window.removeEventListener("keydown", onKey)
  }, [])

  const goTo = (target: Target, text: string) => {
    const current = rangeName(controller.selection.ranges[controller.selection.ranges.length - 1])
    setRecent((items) => [current, ...items.filter((item) => item !== current && item !== text)].slice(0, 10))
    controller.names.goTo(target)
    closeDialog()
  }

  return (
    <TooltipProvider>
      <div className="app-window-frame">
        <TitleBar
          controller={controller}
          title={`${doc.title} - Cells`}
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
            tabs={controller.charts.active() ? TABS_IN_CHART : controller.tables.active() ? TABS_IN_TABLE : TABS}
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
              { kind: "item", label: "Collapse the Ribbon", onSelect: () => ribbonHandle.current?.toggleCollapsed() },
            ]}
          />
        ) : null}
        {controller.showFormulaBar ? (
          <FormulaBar
            controller={controller}
            dragging={false}
            onFocusGrid={focusGrid}
            onInsertFunction={() => openDialog("insertFunction")}
          />
        ) : null}
        <main className="sheet-area">
          <GridView controller={controller} onContextMenu={setMenu} />
        </main>
        <SheetTabs controller={controller} onDialog={openDialog} onFocusGrid={focusGrid} />
        <StatusBar controller={controller} onZoomDialog={() => openDialog("zoom")} />
        {menu ? (
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
        {dialog?.kind === "selectData" ? <SelectDataDialog controller={controller} onClose={closeDialog} /> : null}
        {dialog?.kind === "pageSetup" ? (
          <PageSetupDialog controller={controller} tab={dialog.tab} fileName={doc.fileName} onClose={closeDialog} />
        ) : null}
        {dialog?.kind === "createTable" ? (
          <CreateTableDialog controller={controller} init={dialog.init} onClose={closeDialog} />
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
        {dialog?.kind === "options" ? (
          <OptionsDialog
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
