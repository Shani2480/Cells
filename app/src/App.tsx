import { rangeName, type Range } from "@cells/core"
import { useCallback, useEffect, useMemo, useState } from "react"
import { columnChars, charsToPixels, pointsToPixels, rowPoints } from "@/grid/autofit"
import { GridView, type GridContextMenu } from "@/grid/GridView"
import { usePlatform } from "@/platform/context"
import { SheetController } from "@/sheet/controller"
import { ContextMenu, menuIcons, type MenuItem } from "@/shell/ContextMenu"
import { FormulaBar } from "@/shell/FormulaBar"
import { GoToDialog } from "@/shell/GoToDialog"
import { InsertDeleteDialog } from "@/shell/InsertDeleteDialog"
import { SizeDialog } from "@/shell/SizeDialog"
import { StatusBar } from "@/shell/StatusBar"
import { TitleBar } from "@/shell/TitleBar"
import { loadDemo } from "./demo"

type DialogState =
  | { kind: "goto" }
  | { kind: "insert" | "delete" }
  | { kind: "size"; axis: "rows" | "cols"; indexes: number[]; initial: number | null }
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
  const controller = useMemo(() => given ?? new SheetController(), [given])
  const platform = usePlatform()
  const [dialog, setDialog] = useState<DialogState>(null)
  const [menu, setMenu] = useState<GridContextMenu | null>(null)
  const [recent, setRecent] = useState<string[]>([])
  useSystemColorScheme()

  useEffect(() => {
    if (!platform.window.controllable) document.documentElement.classList.add("is-browser")
    loadDemo(controller, new URLSearchParams(window.location.search).get("demo"))
  }, [controller, platform])

  useEffect(() => {
    controller.shell = {
      goTo: () => setDialog({ kind: "goto" }),
      insertDelete: ({ kind }) => setDialog({ kind }),
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
      { kind: "separator" },
    ]
    if (target === "cells") {
      items.push(
        { kind: "item", label: "Insert...", onSelect: () => controller.insertOrDelete("insert") },
        { kind: "item", label: "Delete...", onSelect: () => controller.insertOrDelete("delete") },
        { kind: "item", label: "Clear Contents", onSelect: () => controller.clearContents() },
      )
    } else {
      const axis = target
      items.push(
        { kind: "item", label: "Insert", onSelect: () => controller.insertLines(axis) },
        { kind: "item", label: "Delete", onSelect: () => controller.deleteLines(axis) },
        { kind: "item", label: "Clear Contents", onSelect: () => controller.clearContents() },
        { kind: "separator" },
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

  const goTo = (range: Range, text: string) => {
    const current = rangeName(controller.selection.ranges[controller.selection.ranges.length - 1])
    setRecent((items) => [current, ...items.filter((item) => item !== current && item !== text)].slice(0, 10))
    controller.goTo(range)
    closeDialog()
  }

  return (
    <div className="app-window-frame">
      <TitleBar controller={controller} title="Book1 - Cells" />
      <FormulaBar controller={controller} dragging={false} onFocusGrid={focusGrid} />
      <main className="sheet-area">
        <GridView controller={controller} onContextMenu={setMenu} />
      </main>
      <StatusBar controller={controller} />
      {menu ? (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          items={menuItems(menu.target)}
          onClose={() => (setMenu(null), focusGrid())}
        />
      ) : null}
      {dialog?.kind === "goto" ? <GoToDialog recent={recent} onGo={goTo} onClose={closeDialog} /> : null}
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
    </div>
  )
}
