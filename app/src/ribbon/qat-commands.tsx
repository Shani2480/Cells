import { ArrowDownAZ, ArrowUpAZ, FilePlus2, FolderOpen, Printer, Redo2, Save, Undo2 } from "lucide-react"
import type { Command } from "./registry"

const icon = { size: 16, strokeWidth: 1.75 }

/** Commands that live on the Quick Access Toolbar (and Tell Me) rather than on a tab. */
export function qatCommands(): Command[] {
  return [
    {
      id: "Undo",
      label: "Undo",
      labelFor: (env) => (env.c.wb.undoLabel ? `Undo ${env.c.wb.undoLabel}` : "Undo"),
      shortcut: "Ctrl+Z",
      icon: <Undo2 {...icon} />,
      run: (env) => env.c.undo(),
      enabled: (env) => env.c.wb.canUndo,
    },
    {
      id: "Redo",
      label: "Redo",
      labelFor: (env) => (env.c.wb.redoLabel ? `Redo ${env.c.wb.redoLabel}` : "Redo"),
      shortcut: "Ctrl+Y",
      icon: <Redo2 {...icon} />,
      run: (env) => env.c.redo(),
      enabled: (env) => env.c.wb.canRedo,
    },
    // Files arrive in Phase 4: listed so the QAT menu matches Excel's, disabled until then.
    { id: "FileNew", label: "New", shortcut: "Ctrl+N", icon: <FilePlus2 {...icon} /> },
    { id: "FileOpen", label: "Open", shortcut: "Ctrl+O", icon: <FolderOpen {...icon} /> },
    { id: "FileSave", label: "Save", shortcut: "Ctrl+S", icon: <Save {...icon} /> },
    { id: "FilePrintQuick", label: "Quick Print", icon: <Printer {...icon} /> },
    { id: "SortAscendingExcel", label: "Sort Ascending", icon: <ArrowUpAZ {...icon} /> },
    { id: "SortDescendingExcel", label: "Sort Descending", icon: <ArrowDownAZ {...icon} /> },
  ]
}

/** Excel's Customize Quick Access Toolbar menu, in its order. */
export const QAT_CHOICES = [
  "FileNew",
  "FileOpen",
  "FileSave",
  "FilePrintQuick",
  "Spelling",
  "Undo",
  "Redo",
  "SortAscendingExcel",
  "SortDescendingExcel",
]

export const QAT_DEFAULT = ["Undo", "Redo"]
const QAT_KEY = "cells.qat"

export function loadQat(): string[] {
  try {
    const saved = JSON.parse(localStorage.getItem(QAT_KEY) ?? "null") as unknown
    if (Array.isArray(saved) && saved.every((id) => typeof id === "string")) return saved as string[]
  } catch {
    // Storage blocked: Excel's default toolbar.
  }
  return QAT_DEFAULT
}

export function saveQat(ids: string[]) {
  try {
    localStorage.setItem(QAT_KEY, JSON.stringify(ids))
  } catch {
    // Kept for this session only.
  }
}
