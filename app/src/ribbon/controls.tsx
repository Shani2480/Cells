import { useEffect, useRef, useState, type ReactNode } from "react"
import { cn } from "@/lib/cn"
import { Popover } from "@/ui/popover"
import { Tip } from "@/ui/tooltip"
import { useRibbon } from "./context"
import { PendingIcon } from "./icons"
import { MenuList } from "./menu"
import { NOT_BUILT, iconOf, shortcutText, type Command } from "./registry"

function Chevron() {
  return (
    <svg className="ribbon-chevron" viewBox="0 0 8 8" aria-hidden>
      <path d="M1.6 2.8 4 5.2 6.4 2.8" fill="none" stroke="currentColor" strokeWidth="1" strokeLinecap="round" />
    </svg>
  )
}

/** The KeyTip badge shown on a control while KeyTips are on. */
export function KeyTip({ id }: { id: string }) {
  const { keytips } = useRibbon()
  const tip = keytips?.get(id)
  return tip ? <span className="keytip">{tip}</span> : null
}

function useCommand(id: string, label?: string) {
  const { env, registry, openMenu, setOpenMenu, afterRun } = useRibbon()
  const command: Command = registry.get(id) ?? { id, label: label ?? id }
  const built = registry.isBuilt(id)
  const enabled = registry.isEnabled(id, env)
  const pressed = command.pressed?.(env) ?? false
  const entries = command.menu ? (typeof command.menu === "function" ? command.menu(env) : command.menu) : null
  const open = openMenu === id
  const run = () => {
    if (!enabled || !command.run) return
    command.run(env)
    afterRun()
  }
  // A tab names its button without the menu's dots ("Protect Sheet" on Review, "Protect Sheet..." in a menu).
  const state = command.labelFor?.(env)
  const dynamic = state && label && !label.endsWith("...") ? state.replace(/\.\.\.$/, "") : state
  const tip = {
    label: dynamic ?? (label || command.label),
    description: command.description,
    shortcut: shortcutText(command.shortcut),
    note: built ? undefined : NOT_BUILT,
  }
  return {
    command,
    /** The name to show: the state's own, else the tab's, else the command's. */
    name: dynamic ?? (label || command.label),
    built,
    enabled,
    pressed,
    entries,
    open,
    setOpen: (next: boolean) => setOpenMenu(next ? id : null),
    run,
    tip,
    env,
  }
}

/** Keeps the sheet's focus: pressing a ribbon button must not take the caret out of the grid. */
const keepFocus = (event: React.MouseEvent) => event.preventDefault()

function MenuPopover({ id, trigger, className }: { id: string; trigger: React.ReactElement; className?: string }) {
  const c = useCommand(id)
  if (!c.entries) return trigger
  return (
    <Popover open={c.open} onOpenChange={c.setOpen} trigger={trigger} className={className}>
      <MenuList entries={c.entries} close={() => c.setOpen(false)} />
    </Popover>
  )
}

/** A 32 px button with its name below; a split button when it both runs and has a menu. */
export function LargeButton({ id, label }: { id: string; label?: string }) {
  const c = useCommand(id, label)
  const name = c.name
  const icon = iconOf(c.command, c.env, true) ?? <PendingIcon size={32} />
  const disabled = !c.enabled
  if (c.entries && c.command.run) {
    return (
      <div className={cn("ribbon-large is-split", c.open && "is-open", disabled && "is-disabled")} data-command={id}>
        <Tip {...c.tip}>
          <button
            type="button"
            className="ribbon-large-top"
            aria-label={name}
            aria-pressed={c.command.pressed ? c.pressed : undefined}
            aria-disabled={disabled || undefined}
            onMouseDown={keepFocus}
            onClick={c.run}
          >
            <span className="ribbon-large-icon">{icon}</span>
          </button>
        </Tip>
        <MenuPopover
          id={id}
          trigger={
            <button
              type="button"
              className="ribbon-large-bottom"
              aria-label={`${name} options`}
              aria-haspopup="menu"
              aria-disabled={disabled || undefined}
              onMouseDown={keepFocus}
            >
              <LargeLabel text={name} menu />
            </button>
          }
        />
        <KeyTip id={id} />
      </div>
    )
  }
  const button = (
    <button
      type="button"
      className={cn("ribbon-large", c.open && "is-open")}
      aria-label={name}
      aria-pressed={c.command.pressed ? c.pressed : undefined}
      aria-haspopup={c.entries ? "menu" : undefined}
      aria-disabled={disabled || undefined}
      data-command={id}
      onMouseDown={keepFocus}
      onClick={c.entries ? undefined : c.run}
    >
      <span className="ribbon-large-icon">{icon}</span>
      <LargeLabel text={name} menu={!!c.entries} />
      <KeyTip id={id} />
    </button>
  )
  return c.entries && !disabled ? (
    <MenuPopover id={id} trigger={<Tip {...c.tip}>{button}</Tip>} />
  ) : (
    <Tip {...c.tip}>{button}</Tip>
  )
}

/** Large labels break onto two lines at the last space, the chevron after the second line (as Excel). */
function LargeLabel({ text, menu }: { text: string; menu?: boolean }) {
  const at = text.lastIndexOf(" ")
  const first = at > 0 ? text.slice(0, at) : text
  const second = at > 0 ? text.slice(at + 1) : ""
  return (
    <span className="ribbon-large-label">
      <span>{first}</span>
      <span>
        {second}
        {menu ? <Chevron /> : null}
      </span>
    </span>
  )
}

/** A small button: picture, and its name when there is room; split when it both runs and has a menu. */
export function SmallButton({ id, label, showLabel }: { id: string; label?: string; showLabel: boolean }) {
  const c = useCommand(id, label)
  const name = label === "" ? "" : c.name
  const icon = iconOf(c.command, c.env) ?? <PendingIcon />
  const disabled = !c.enabled
  const body = (
    <>
      <span className="ribbon-small-icon">{icon}</span>
      {showLabel && name ? <span className="ribbon-small-label">{name}</span> : null}
    </>
  )
  if (c.entries && c.command.run) {
    return (
      <div className={cn("ribbon-small is-split", c.open && "is-open", disabled && "is-disabled")} data-command={id}>
        <Tip {...c.tip}>
          <button
            type="button"
            className="ribbon-small-main"
            aria-label={c.tip.label}
            aria-pressed={c.command.pressed ? c.pressed : undefined}
            aria-disabled={disabled || undefined}
            onMouseDown={keepFocus}
            onClick={c.run}
          >
            {body}
          </button>
        </Tip>
        <MenuPopover
          id={id}
          trigger={
            <button
              type="button"
              className="ribbon-small-arrow"
              aria-label={`${c.tip.label} options`}
              aria-haspopup="menu"
              aria-disabled={disabled || undefined}
              onMouseDown={keepFocus}
            >
              <Chevron />
            </button>
          }
        />
        <KeyTip id={id} />
      </div>
    )
  }
  const button = (
    <button
      type="button"
      className={cn("ribbon-small", c.open && "is-open")}
      aria-label={c.tip.label}
      aria-pressed={c.command.pressed ? c.pressed : undefined}
      aria-haspopup={c.entries ? "menu" : undefined}
      aria-disabled={disabled || undefined}
      data-command={id}
      onMouseDown={keepFocus}
      onClick={c.entries ? undefined : c.run}
    >
      {body}
      {c.entries ? <Chevron /> : null}
      <KeyTip id={id} />
    </button>
  )
  return c.entries && !disabled ? (
    <MenuPopover id={id} trigger={<Tip {...c.tip}>{button}</Tip>} />
  ) : (
    <Tip {...c.tip}>{button}</Tip>
  )
}

/** A checkbox with its name (View › Gridlines). */
export function CheckControl({ id, label }: { id: string; label?: string }) {
  const c = useCommand(id, label)
  return (
    <Tip {...c.tip}>
      <button
        type="button"
        role="checkbox"
        className="ribbon-check"
        aria-checked={c.pressed}
        aria-disabled={!c.enabled || undefined}
        data-command={id}
        onMouseDown={keepFocus}
        onClick={c.run}
      >
        <span className="ribbon-check-box">
          {c.pressed ? (
            <svg viewBox="0 0 12 12" aria-hidden>
              <path d="m2.5 6.2 2.3 2.3 4.7-5" fill="none" stroke="currentColor" strokeWidth="1.5" />
            </svg>
          ) : null}
        </span>
        <span>{label || c.command.label}</span>
        <KeyTip id={id} />
      </button>
    </Tip>
  )
}

/** A combo box (Font, Font Size, Number Format): type a value and press Enter, or pick from the list. */
export function FieldControl({ id }: { id: string }) {
  const c = useCommand(id)
  const field = c.command.field
  const current = field ? field.value(c.env) : ""
  const [text, setText] = useState(current)
  const [editing, setEditing] = useState(false)
  const input = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (!editing) setText(current)
  }, [current, editing])
  if (!field) return null
  const options = field.options?.(c.env) ?? []
  const commit = (value: string) => {
    field.commit(c.env, value)
    setEditing(false)
    c.setOpen(false)
    c.env.c.changed()
  }
  return (
    <div className={cn("ribbon-field", c.open && "is-open")} style={{ width: field.width }} data-command={id}>
      <Tip {...c.tip}>
        <input
          ref={input}
          className="ribbon-field-input"
          aria-label={c.command.label}
          value={text}
          spellCheck={false}
          onFocus={(event) => {
            setEditing(true)
            event.currentTarget.select()
          }}
          onBlur={() => {
            setEditing(false)
            setText(current)
          }}
          onChange={(event) => setText(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault()
              commit(text)
              input.current?.blur()
              focusSheet()
            } else if (event.key === "Escape") {
              event.preventDefault()
              setText(current)
              input.current?.blur()
              focusSheet()
            }
          }}
        />
      </Tip>
      {options.length ? (
        <Popover
          open={c.open}
          onOpenChange={c.setOpen}
          className="field-list"
          trigger={
            <button
              type="button"
              className="ribbon-field-arrow"
              aria-label={`${c.command.label} list`}
              aria-haspopup="listbox"
              onMouseDown={keepFocus}
            >
              <Chevron />
            </button>
          }
        >
          <div role="listbox" className="field-options">
            {options.map((o) => (
              <button
                key={o.value}
                type="button"
                role="option"
                aria-selected={o.label === current}
                className="field-option"
                onClick={() => {
                  commit(o.value)
                  focusSheet()
                }}
              >
                {o.preview ?? o.label}
              </button>
            ))}
          </div>
        </Popover>
      ) : null}
      <KeyTip id={id} />
    </div>
  )
}

/** The active window's hidden editor textarea (with Arrange All, each window's grid has one). */
export function sheetEditor(): HTMLTextAreaElement | null {
  return (
    document.querySelector<HTMLTextAreaElement>(".win-panel.is-active .cell-editor") ??
    document.querySelector<HTMLTextAreaElement>(".cell-editor")
  )
}

/** Puts the keyboard back on the sheet (its hidden editor textarea). */
export function focusSheet() {
  requestAnimationFrame(() => sheetEditor()?.focus({ preventScroll: true }))
}

/** After a ribbon command: the keyboard back where typing was (a shape's text being typed, else the sheet). */
export function focusAfterCommand() {
  requestAnimationFrame(() => {
    // A command that opened a dialog leaves the focus in it (the menu it was run from is a popup too, but closing).
    const dialog = document.activeElement?.closest("[role=dialog], [role=alertdialog]")
    if (dialog && !dialog.closest(".popover")) return
    const shapeText = document.querySelector<HTMLTextAreaElement>(".object-text-editor")
    ;(shapeText ?? sheetEditor())?.focus({ preventScroll: true })
  })
}

export function GroupLauncher({ id, groupLabel }: { id: string; groupLabel: string }) {
  const c = useCommand(id)
  return (
    <Tip {...c.tip}>
      <button
        type="button"
        className="ribbon-launcher"
        aria-label={c.tip.label || groupLabel}
        aria-disabled={!c.enabled || undefined}
        onMouseDown={keepFocus}
        onClick={c.run}
      >
        <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden>
          <path d="M2 2v6h6M5 5l3.5 3.5M8.5 5.5v3h-3" fill="none" stroke="currentColor" strokeWidth="1" />
        </svg>
        <KeyTip id={id} />
      </button>
    </Tip>
  )
}

export function RibbonLabel({ children }: { children: ReactNode }) {
  return <span className="ribbon-text">{children}</span>
}
