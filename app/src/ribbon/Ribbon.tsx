import { ChevronDown, ChevronUp } from "lucide-react"
import {
  Fragment,
  useCallback,
  useContext,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from "react"
import { cn } from "@/lib/cn"
import { Popover } from "@/ui/popover"
import { Tip } from "@/ui/tooltip"
import { RibbonCtx, type RibbonState } from "./context"
import {
  CheckControl,
  FieldControl,
  GroupLauncher,
  LargeButton,
  RibbonLabel,
  SmallButton,
  focusAfterCommand,
  focusSheet,
} from "./controls"
import { fitGroups, type GroupWidths, type Level } from "./fit"
import { GroupIcon } from "./icons"
import { keytipsFor } from "./keytips"
import type { CommandRegistry, RibbonEnv } from "./registry"
import { controlIds, type Column, type Control, type Group, type Tab } from "./tabs"

const STORE_KEY = "cells.ribbon.collapsed"

function readCollapsed(): boolean {
  try {
    return localStorage.getItem(STORE_KEY) === "1"
  } catch {
    return false
  }
}

function writeCollapsed(value: boolean) {
  try {
    localStorage.setItem(STORE_KEY, value ? "1" : "0")
  } catch {
    // Storage can be off (private mode); the choice then lasts for this session only.
  }
}

/** KeyTips: off, showing the tabs' letters, or showing a tab's controls' letters (with what was typed so far). */
type KeyTipState = null | { stage: "tabs"; typed?: string } | { stage: "tab"; typed: string }

export type RibbonHandle = {
  toggleCollapsed(): void
  /** Search results that open a menu or a box: shows the tab with the control and opens it. False if no tab has it. */
  show(id: string): boolean
  /** Selects a tab, now or as soon as it shows (a contextual tab after the object it belongs to is made). */
  showTab(id: string): void
}

export function Ribbon({
  env,
  registry,
  tabs,
  onFile,
  qatKeytips,
  handle,
}: {
  env: RibbonEnv
  registry: CommandRegistry
  tabs: Tab[]
  /** The File tab opens the backstage view. */
  onFile: () => void
  /** KeyTips for the Quick Access Toolbar's buttons ("1" → run). */
  qatKeytips?: { key: string; run: () => void }[]
  /** Lets the app toggle the ribbon (Ctrl+F1 from the sheet's keys). */
  handle?: (handle: RibbonHandle) => void
}) {
  // Pressed states and boxes follow the sheet (the active cell's format), so redraw on every change.
  useSyncExternalStore(env.c.subscribe, env.c.getVersion)
  const [tabId, setTabId] = useState(tabs[0].id)
  const [collapsed, setCollapsed] = useState(readCollapsed)
  /** A collapsed ribbon shows a tab's commands over the sheet while one is picked. */
  const [peek, setPeek] = useState(false)
  const [openMenu, setOpenMenu] = useState<string | null>(null)
  // A gallery's live preview ends with its menu, however the menu closes (only then: Quick Analysis previews too).
  const hadMenu = useRef(false)
  useEffect(() => {
    if (openMenu === null && hadMenu.current && env.c.conditional.preview) env.c.conditional.setPreview(null)
    hadMenu.current = openMenu !== null
  }, [openMenu, env])
  const [kt, setKt] = useState<KeyTipState>(null)
  const tab = tabs.find((t) => t.id === tabId) ?? tabs[0]
  // A contextual tab that goes (the cell left the table) hands back to Home, as in Excel; it isn't picked again later.
  useEffect(() => {
    if (!tabs.some((t) => t.id === tabId)) setTabId(tabs[0].id)
  }, [tabs, tabId])
  const rootRef = useRef<HTMLDivElement>(null)

  const toggleCollapsed = useCallback(() => {
    setCollapsed((value) => {
      writeCollapsed(!value)
      return !value
    })
    setPeek(false)
  }, [])

  const show = useCallback(
    (id: string) => {
      const home = tabs.find((t) => controlIds(t).includes(id))
      if (!home) return false
      setTabId(home.id)
      if (collapsed) setPeek(true)
      // After the tab has drawn its controls.
      requestAnimationFrame(() => {
        const el = rootRef.current?.querySelector<HTMLElement>(`.ribbon-groups [data-command="${CSS.escape(id)}"]`)
        const command = registry.get(id)
        if (command?.field) el?.querySelector<HTMLInputElement>("input")?.focus()
        else if (command?.menu) setOpenMenu(id)
        else el?.querySelector<HTMLElement>("button")?.focus()
      })
      return true
    },
    [tabs, collapsed, registry],
  )

  const [wantTab, setWantTab] = useState<string | null>(null)
  useEffect(() => {
    if (wantTab && tabs.some((t) => t.id === wantTab)) {
      setTabId(wantTab)
      setWantTab(null)
    }
  }, [wantTab, tabs])

  useEffect(() => {
    handle?.({ toggleCollapsed, show, showTab: setWantTab })
  }, [handle, toggleCollapsed, show])

  const afterRun = useCallback(() => {
    setPeek(false)
    setKt(null)
    focusAfterCommand()
  }, [])

  const keytips = useMemo(() => (kt?.stage === "tab" ? keytipsFor(tab, registry, kt.typed) : null), [kt, tab, registry])

  const state: RibbonState = { env, registry, openMenu, setOpenMenu, keytips, afterRun }

  // ── KeyTips: Alt (or F10) shows letters; a letter picks a tab, then a control. ──
  useEffect(() => {
    let altAlone = false
    const onDown = (event: KeyboardEvent) => {
      if (event.key === "Alt") {
        altAlone = !event.repeat
        return
      }
      altAlone = false
      if (kt) {
        handleKeyTip(event)
        return
      }
      // Alt+H and friends go straight to a tab's KeyTips (not while typing in a cell or field).
      if (event.altKey && !event.ctrlKey && !event.metaKey && !env.c.edit && !isTextField(event.target)) {
        const letter = event.key.toUpperCase()
        const target = tabs.find((t) => t.keytip === letter)
        if (target) {
          event.preventDefault()
          event.stopPropagation()
          setTabId(target.id)
          if (collapsed) setPeek(true)
          setKt({ stage: "tab", typed: "" })
        } else if (letter === "F") {
          event.preventDefault()
          onFile()
        }
      }
      if (event.key === "F10" && !event.shiftKey) {
        event.preventDefault()
        setKt({ stage: "tabs" })
      }
    }
    const onUp = (event: KeyboardEvent) => {
      if (event.key === "Alt" && altAlone) {
        altAlone = false
        if (isTextField(event.target) || env.c.edit) return
        event.preventDefault()
        setKt((current) => (current ? null : { stage: "tabs" }))
      }
    }
    const onPointer = () => {
      altAlone = false
      setKt(null)
    }
    const handleKeyTip = (event: KeyboardEvent) => {
      if (event.ctrlKey || event.metaKey) return
      event.preventDefault()
      event.stopPropagation()
      if (event.key === "Escape") {
        setKt(kt?.stage === "tab" ? { stage: "tabs" } : null)
        if (kt?.stage === "tabs") focusSheet()
        return
      }
      if (event.key.length !== 1) return
      const letter = event.key.toUpperCase()
      if (kt?.stage === "tabs") {
        // Contextual tabs have two letters (Table Design: J, T).
        const typed = (kt.typed ?? "") + letter
        const target = tabs.find((t) => t.keytip === typed)
        if (!target && tabs.some((t) => t.keytip.startsWith(typed))) {
          setKt({ stage: "tabs", typed })
          return
        }
        if (target) {
          setTabId(target.id)
          if (collapsed) setPeek(true)
          setKt({ stage: "tab", typed: "" })
        } else if (letter === "F") {
          setKt(null)
          onFile()
        } else {
          const qat = qatKeytips?.find((q) => q.key === letter)
          if (qat) {
            setKt(null)
            qat.run()
          }
        }
        return
      }
      if (kt?.stage === "tab") {
        const typed = kt.typed + letter
        const all = keytipsFor(tab, registry, "")
        const hit = [...all].find(([, key]) => key === typed)
        if (hit) {
          activate(hit[0])
          return
        }
        if ([...all.values()].some((key) => key.startsWith(typed))) setKt({ stage: "tab", typed })
      }
    }
    const activate = (id: string) => {
      setKt(null)
      const el = rootRef.current?.querySelector<HTMLElement>(`.ribbon-groups [data-command="${CSS.escape(id)}"]`)
      const command = registry.get(id)
      if (!command || !registry.isEnabled(id, env)) return
      if (command.field) {
        el?.querySelector<HTMLInputElement>("input")?.focus()
        return
      }
      if (command.menu && !command.run) {
        setOpenMenu(id)
        return
      }
      // A split button's KeyTip opens its menu in Excel only for menu-first controls; others run.
      command.run?.(env)
      afterRun()
    }
    window.addEventListener("keydown", onDown, true)
    window.addEventListener("keyup", onUp, true)
    window.addEventListener("pointerdown", onPointer, true)
    return () => {
      window.removeEventListener("keydown", onDown, true)
      window.removeEventListener("keyup", onUp, true)
      window.removeEventListener("pointerdown", onPointer, true)
    }
  }, [kt, tabs, tab, registry, env, collapsed, onFile, qatKeytips, afterRun])

  // A peeking ribbon closes on a click outside it.
  useEffect(() => {
    if (!peek) return
    const onDown = (event: PointerEvent) => {
      const target = event.target as Element
      if (rootRef.current?.contains(target) || target.closest?.(".popover-positioner, .tip-positioner")) return
      setPeek(false)
    }
    window.addEventListener("pointerdown", onDown, true)
    return () => window.removeEventListener("pointerdown", onDown, true)
  }, [peek])

  const showPanel = !collapsed || peek
  return (
    <RibbonCtx.Provider value={state}>
      <div ref={rootRef} className={cn("ribbon", collapsed && "is-collapsed", kt && "is-keytips")}>
        <div className="ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
          <button
            type="button"
            className="ribbon-file"
            onMouseDown={(event) => event.preventDefault()}
            onClick={onFile}
          >
            File
            {kt?.stage === "tabs" ? <span className="keytip">F</span> : null}
          </button>
          {tabs.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={t.id === tab.id && showPanel}
              className={cn("ribbon-tab", t.contextual && "is-contextual")}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                if (collapsed) setPeek(!(peek && t.id === tab.id))
                setTabId(t.id)
              }}
              onDoubleClick={toggleCollapsed}
            >
              {t.label}
              {kt?.stage === "tabs" ? <span className="keytip">{t.keytip}</span> : null}
            </button>
          ))}
        </div>
        {showPanel ? (
          <div className={cn("ribbon-panel", collapsed && "is-peek")} role="tabpanel" aria-label={tab.label}>
            <FittedGroups key={tab.id} tab={tab} />
            <Tip
              label={collapsed ? "Pin the Ribbon" : "Collapse the Ribbon"}
              description={
                collapsed
                  ? "Keep the ribbon open all the time."
                  : "Need more space on your screen? Show only the tab names; click a tab to see its commands."
              }
              shortcut="Ctrl+F1"
            >
              <button
                type="button"
                className="ribbon-toggle"
                aria-label={collapsed ? "Pin the Ribbon" : "Collapse the Ribbon"}
                onMouseDown={(event) => event.preventDefault()}
                onClick={toggleCollapsed}
              >
                {collapsed ? <ChevronDown size={14} strokeWidth={1.75} /> : <ChevronUp size={14} strokeWidth={1.75} />}
              </button>
            </Tip>
          </div>
        ) : null}
      </div>
    </RibbonCtx.Provider>
  )
}

function isTextField(target: EventTarget | null) {
  return (
    target instanceof HTMLInputElement ||
    (target instanceof HTMLTextAreaElement && !target.classList.contains("cell-editor")) ||
    (target instanceof HTMLElement && target.isContentEditable)
  )
}

/** The tab's groups at the levels that fit the panel's width (measured from a hidden copy at every level). */
function FittedGroups({ tab }: { tab: Tab }) {
  const ribbon = useContext(RibbonCtx)!
  const box = useRef<HTMLDivElement>(null)
  const measurer = useRef<HTMLDivElement>(null)
  const [widths, setWidths] = useState<GroupWidths[] | null>(null)
  const [available, setAvailable] = useState(Infinity)

  useLayoutEffect(() => {
    const measure = () => {
      const rows = measurer.current?.querySelectorAll<HTMLElement>("[data-measure-row]")
      if (!rows) return
      const out: GroupWidths[] = tab.groups.map(() => [0, 0, 0])
      rows.forEach((row) => {
        const level = Number(row.dataset.measureRow)
        row.querySelectorAll<HTMLElement>(":scope > .ribbon-group").forEach((g, i) => {
          out[i][level] = g.getBoundingClientRect().width
        })
      })
      setWidths(out)
    }
    measure()
    void document.fonts?.ready.then(measure)
  }, [tab])

  useLayoutEffect(() => {
    const el = box.current
    if (!el) return
    const update = () => setAvailable(el.clientWidth)
    update()
    const observer = new ResizeObserver(update)
    observer.observe(el)
    return () => observer.disconnect()
  }, [])

  const levels: Level[] = widths ? fitGroups(widths, available - 4) : tab.groups.map(() => 0)
  return (
    <>
      <div ref={box} className="ribbon-groups">
        {tab.groups.map((group, i) => (
          <GroupView key={group.id} group={group} level={levels[i]} />
        ))}
      </div>
      {/* The measuring copy never opens menus or shows KeyTips (its popups would portal out, visible). */}
      <RibbonCtx.Provider value={{ ...ribbon, openMenu: null, setOpenMenu: () => undefined, keytips: null }}>
        <div ref={measurer} className="ribbon-measure" aria-hidden inert>
          {([0, 1, 2] as const).map((level) => (
            <div key={level} data-measure-row={level} className="ribbon-measure-row">
              {tab.groups.map((group) => (
                <GroupView key={group.id} group={group} level={level} />
              ))}
            </div>
          ))}
        </div>
      </RibbonCtx.Provider>
    </>
  )
}

function GroupView({ group, level }: { group: Group; level: Level }) {
  if (level === 2) return <CollapsedGroup group={group} />
  return (
    <div className="ribbon-group" data-group={group.id} data-level={level}>
      <div className="ribbon-group-body">{renderColumns(group.columns, level)}</div>
      <div className="ribbon-caption">
        <span>{group.label}</span>
        {group.launcher ? <GroupLauncher id={group.launcher} groupLabel={group.label} /> : null}
      </div>
    </div>
  )
}

/** Compact: runs of large buttons stack three to a column as small ones; small buttons drop their names. */
function renderColumns(columns: Column[], level: Level): ReactNode {
  if (level === 0) return columns.map((col, i) => <ColumnView key={i} column={col} compact={false} />)
  const out: ReactNode[] = []
  let run: Control[] = []
  const flush = () => {
    if (!run.length) return
    for (let i = 0; i < run.length; i += 3) {
      const chunk = run.slice(i, i + 3)
      out.push(
        <div key={`r${out.length}`} className="ribbon-stack">
          {chunk.map((control) => (
            <div key={(control as { id: string }).id} className="ribbon-row">
              <ControlView control={{ ...control, kind: "small" } as Control} showLabel />
            </div>
          ))}
        </div>,
      )
    }
    run = []
  }
  columns.forEach((col, i) => {
    // A lone large button stays large; two or more in a row stack up.
    if (col.kind === "large" && (columns[i - 1]?.kind === "large" || columns[i + 1]?.kind === "large")) {
      run.push(col.control)
      return
    }
    flush()
    out.push(<ColumnView key={`c${i}`} column={col} compact />)
  })
  flush()
  return out
}

function ColumnView({ column, compact }: { column: Column; compact: boolean }) {
  if (column.kind === "large") return <ControlView control={column.control} showLabel />
  return (
    <div className={cn("ribbon-stack", column.rows.length < 3 && "is-two-row")}>
      {column.rows.map((row, i) => (
        <div key={i} className="ribbon-row">
          {row.map((control, j) => (
            <Fragment key={j}>
              <ControlView control={control} showLabel={row.length === 1 && !compact} />
            </Fragment>
          ))}
        </div>
      ))}
    </div>
  )
}

function ControlView({ control, showLabel }: { control: Control; showLabel: boolean }) {
  switch (control.kind) {
    case "large":
      return <LargeButton id={control.id} label={control.label} />
    case "small":
      return <SmallButton id={control.id} label={control.label} showLabel={showLabel && !control.iconOnly} />
    case "field":
      return <FieldControl id={control.id} />
    case "check":
      return <CheckControl id={control.id} label={control.label} />
    case "label":
      return <RibbonLabel>{control.text}</RibbonLabel>
  }
}

/** A group too wide for the window: one button with the group's name that opens the whole group. */
function CollapsedGroup({ group }: { group: Group }) {
  const [open, setOpen] = useState(false)
  return (
    <div className="ribbon-group is-collapsed" data-group={group.id} data-level={2}>
      <Popover
        open={open}
        onOpenChange={setOpen}
        className="ribbon-group-popup"
        trigger={
          <button
            type="button"
            className={cn("ribbon-large", open && "is-open")}
            aria-label={group.label}
            aria-haspopup="dialog"
            data-command={`group:${group.id}`}
            onMouseDown={(event) => event.preventDefault()}
          >
            <span className="ribbon-large-icon">
              <GroupIcon />
            </span>
            <span className="ribbon-large-label">
              <span>{group.label}</span>
              <span>
                <svg className="ribbon-chevron" viewBox="0 0 8 8" aria-hidden>
                  <path d="M1.6 2.8 4 5.2 6.4 2.8" fill="none" stroke="currentColor" strokeWidth="1" />
                </svg>
              </span>
            </span>
          </button>
        }
      >
        <GroupView group={group} level={0} />
      </Popover>
    </div>
  )
}
