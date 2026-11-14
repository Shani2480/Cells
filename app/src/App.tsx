import { rangeName } from "@cells/core"
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react"
import { columnChars, charsToPixels, pointsToPixels, rowPoints } from "@/grid/autofit"
import { GridView, type GridContextMenu } from "@/grid/GridView"
import { usePlatform } from "@/platform/context"
import { SheetController } from "@/sheet/controller"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { Ribbon, type RibbonHandle } from "@/ribbon/Ribbon"
import { TABS } from "@/ribbon/tabs"
import { Backstage } from "@/shell/Backstage"
import { whatsNewOnStart } from "@/about/release-notes"
import { Dialog } from "@/shell/Dialog"
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
import { startCalculation } from "./engine-host/host"

const NOT_BUILT = "Not built yet."

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
  | { kind: "cfNew"; start: string }
  | null

/** Follows the system's light/dark setting (a manual choice comes with File > Options in Phase 3). */
function useSystemColorScheme() {
  useEffect(() => {
    const query = window.matchMedia("(prefers-color-scheme: dark)")
    const apply = () => document.documentElement.classList.toggle("dark", query.matches)
    apply()
    query.addEventListener("change", apply)
    return () => query.removeEventListener("change", apply)
  }, [])
}

export function App({ controller: given }: { controller?: SheetController } = {}) {
  const controller = useMemo(() => {
    if (given) return given
    const created = new SheetController()
    // Formulas calculate in a worker; the grid shows the values as they come back.
    startCalculation(created.wb)
    return created
  }, [given])
  const platform = usePlatform()
  const [dialog, setDialog] = useState<DialogState>(null)
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
  useSystemColorScheme()

  useEffect(() => {
    if (!platform.window.controllable) document.documentElement.classList.add("is-browser")
    loadDemo(controller, new URLSearchParams(window.location.search).get("demo"))
  }, [controller, platform])

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
    backstage: (page) => setBackstage(page ?? "account"),
  }

  // Ctrl+F1 shows or hides the ribbon's commands, from anywhere.
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "F1" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault()
        ribbonHandle.current?.toggleCollapsed()
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
          title="Book1 - Cells"
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
            tabs={TABS}
            onFile={() => setBackstage("account")}
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
