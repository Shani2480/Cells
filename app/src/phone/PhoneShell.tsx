import {
  Check,
  ChevronDown,
  Eye,
  ListFilter,
  Menu,
  Plus,
  Redo2,
  Search,
  SquareFunction,
  Type,
  Undo2,
} from "lucide-react"
import { useEffect, useState, useSyncExternalStore, type ReactNode } from "react"
import type { CommandRegistry, RibbonEnv } from "@/ribbon/registry"
import type { Tab } from "@/ribbon/tabs"
import type { SheetController } from "@/sheet/controller"
import { BottomSheet } from "./BottomSheet"
import { CommandSheet } from "./CommandSheet"

/**
 * The window on a phone (shell-choice.ts): a compact top bar (File, the workbook's name, Undo, Redo, Find), the
 * formula bar, the sheet, and a bottom bar with the sheet switcher and the tool buttons that open the command palette
 * on a tab. While a cell is being edited the bottom bar gives way to a row of formula keys, just above the on-screen
 * keyboard. Dialogs, the File view and menus are the full window's own (shown full screen on a phone by the style).
 */

/** The keys over the keyboard while editing: what a phone's keyboard hides behind its symbol pages. */
export const FORMULA_KEYS = ["=", "(", ")", ",", ":", "$", "+", "-", "*", "/", "&", '"'] as const

/** Types `key` at the editor's caret, over any selected text. */
export function typeAtCaret(controller: SheetController, key: string) {
  const edit = controller.edit
  if (!edit) return
  const text = edit.text.slice(0, edit.selStart) + key + edit.text.slice(edit.selEnd)
  controller.setEditText(text, edit.selStart + key.length)
}

/** The tool buttons: a name, an icon and the tab the palette opens on. */
const TOOLS: { tab: string; label: string; icon: ReactNode }[] = [
  { tab: "TabHome", label: "Home", icon: <Type size={20} aria-hidden /> },
  { tab: "TabInsert", label: "Insert", icon: <Plus size={20} aria-hidden /> },
  { tab: "TabFormulas", label: "Formulas", icon: <SquareFunction size={20} aria-hidden /> },
  { tab: "TabData", label: "Data", icon: <ListFilter size={20} aria-hidden /> },
  { tab: "TabView", label: "View", icon: <Eye size={20} aria-hidden /> },
]

export function PhoneShell({
  controller,
  env,
  registry,
  tabs,
  title,
  formulaBar,
  children,
  onFile,
  onFocusGrid,
  onTitle,
  workbooks = 1,
}: {
  controller: SheetController
  env: RibbonEnv
  registry: CommandRegistry
  /** The palette's tabs: the ribbon's, with a table's or chart's own tab while one is selected. */
  tabs: Tab[]
  title: string
  formulaBar: ReactNode
  /** The sheet. */
  children: ReactNode
  onFile: () => void
  onFocusGrid: () => void
  /** Tapping the name shows the open workbooks (where several share the window). */
  onTitle?: () => void
  /** How many workbooks are open, shown beside the name when more than one. */
  workbooks?: number
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const [palette, setPalette] = useState<string | null>(null)
  const [sheets, setSheets] = useState(false)

  // Finger-sized column and row headings while the phone shell is shown.
  useEffect(() => {
    controller.largeHeaders = true
    controller.changed()
    return () => {
      controller.largeHeaders = false
      controller.changed()
    }
  }, [controller])

  const run = (id: string) => {
    if (registry.isEnabled(id, env)) registry.get(id)?.run?.(env)
  }
  const tool = (id: string, label: string, icon: ReactNode) => (
    <button
      type="button"
      className="phone-icon-button"
      aria-label={label}
      title={label}
      disabled={!registry.isEnabled(id, env)}
      onClick={() => {
        run(id)
        onFocusGrid()
      }}
    >
      {icon}
    </button>
  )

  const editing = controller.edit
  const active = controller.wb.sheets[controller.sheetIndex]
  return (
    <div className="phone-window">
      <header className="phone-topbar">
        <button type="button" className="phone-icon-button" aria-label="File" title="File" onClick={onFile}>
          <Menu size={22} aria-hidden />
        </button>
        {onTitle ? (
          <button
            type="button"
            className="phone-title phone-title-switch"
            aria-label="Switch workbooks"
            onClick={onTitle}
          >
            <span className="phone-title-name">{title}</span>
            {workbooks > 1 ? <span className="wb-switch-count">{workbooks}</span> : null}
            <ChevronDown size={16} aria-hidden />
          </button>
        ) : (
          <span className="phone-title">{title}</span>
        )}
        {tool("Undo", "Undo", <Undo2 size={20} />)}
        {tool("Redo", "Redo", <Redo2 size={20} />)}
        {tool("FindDialogExcel", "Find", <Search size={20} />)}
      </header>
      {formulaBar}
      <main className="sheet-area phone-sheet-area">{children}</main>
      {editing ? (
        <div className="phone-keys" role="toolbar" aria-label="Formula keys">
          <button
            type="button"
            className="phone-key is-fx"
            aria-label="Insert Function"
            onPointerDown={(event) => event.preventDefault()}
            onClick={() => env.dialog("insertFunction")}
          >
            <SquareFunction size={18} aria-hidden />
          </button>
          <div className="phone-keys-scroll">
            {FORMULA_KEYS.map((key) => (
              <button
                key={key}
                type="button"
                className="phone-key"
                // Keeps the caret, and the keyboard, in the editor.
                onPointerDown={(event) => event.preventDefault()}
                onClick={() => typeAtCaret(controller, key)}
              >
                {key}
              </button>
            ))}
          </div>
          <button
            type="button"
            className="phone-key is-enter"
            aria-label="Enter"
            onPointerDown={(event) => event.preventDefault()}
            onClick={() => {
              controller.commitEdit()
              onFocusGrid()
            }}
          >
            <Check size={18} aria-hidden />
          </button>
        </div>
      ) : (
        <nav className="phone-bottombar" aria-label="Sheets and tools">
          <button
            type="button"
            className="phone-sheet-button"
            aria-haspopup="dialog"
            aria-label={`Sheet: ${active?.name ?? ""}`}
            onClick={() => setSheets(true)}
          >
            <span className="phone-sheet-name">{active?.name}</span>
            <ChevronDown size={16} aria-hidden />
          </button>
          <div className="phone-tools">
            {TOOLS.filter((t) => tabs.some((tab) => tab.id === t.tab)).map((t) => (
              <button key={t.tab} type="button" className="phone-tool" onClick={() => setPalette(t.tab)}>
                {t.icon}
                <span>{t.label}</span>
              </button>
            ))}
          </div>
        </nav>
      )}
      {palette ? (
        <CommandSheet
          env={env}
          registry={registry}
          tabs={tabs}
          tab={palette}
          onClose={() => {
            setPalette(null)
            onFocusGrid()
          }}
        />
      ) : null}
      {sheets ? (
        <SheetSwitcher
          controller={controller}
          onClose={() => {
            setSheets(false)
            onFocusGrid()
          }}
        />
      ) : null}
    </div>
  )
}

/** The sheet list from the bottom bar: every visible sheet (the active one ticked), and New Sheet. */
export function SheetSwitcher({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const sheets = controller.wb.sheets
  const locked = controller.sheets.isWorkbookProtected()
  return (
    <BottomSheet label="Sheets" onClose={onClose} header={<span className="phone-sheet-title">Sheets</span>}>
      <div role="listbox" aria-label="Sheets">
        {controller.sheets.visible().map((index) => (
          <button
            key={index}
            type="button"
            role="option"
            aria-selected={index === controller.sheetIndex}
            className="phone-row"
            onClick={() => {
              controller.activate(index)
              onClose()
            }}
          >
            <span className="phone-row-icon">
              {index === controller.sheetIndex ? <Check size={18} aria-hidden /> : null}
            </span>
            <span className="phone-row-label">{sheets[index].name}</span>
            {sheets[index].tabColor ? (
              <span className="phone-sheet-color" style={{ background: sheets[index].tabColor }} aria-hidden />
            ) : null}
          </button>
        ))}
      </div>
      <button
        type="button"
        className="phone-row is-action"
        disabled={locked}
        onClick={() => {
          controller.sheets.insert(controller.sheetIndex + 1)
          onClose()
        }}
      >
        <span className="phone-row-icon">
          <Plus size={18} aria-hidden />
        </span>
        <span className="phone-row-label">New Sheet</span>
      </button>
    </BottomSheet>
  )
}
