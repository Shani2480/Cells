import { Check, ChevronDown, ChevronLeft, ChevronRight } from "lucide-react"
import { useState, useSyncExternalStore, type ReactNode } from "react"
import { RibbonCtx, type RibbonState } from "@/ribbon/context"
import { MenuList } from "@/ribbon/menu"
import { iconOf, type CommandRegistry, type RibbonEnv } from "@/ribbon/registry"
import type { Control, Tab } from "@/ribbon/tabs"
import { BottomSheet } from "./BottomSheet"

/**
 * The phone's ribbon, after Excel for phones' command palette: a sheet from the bottom with a tab picker at its top
 * and the tab's commands below, group by group, one finger-sized row each. A dropdown opens its menu in place with a
 * Back row; a split button runs from its row and opens its menu from the › beside it; a field (font, size, number
 * format) shows its value and lists its choices in place. The commands are the ribbon's own (the same registry), and
 * the ones Cells doesn't have yet are left out rather than shown dead.
 */

/** One row of the palette: a control of the tab, when it does something. */
export type PaletteRow = { id: string; label?: string; kind: "command" | "check" | "field" }

export type PaletteGroup = { id: string; label: string; rows: PaletteRow[] }

/** A tab's groups as palette rows: built controls only, launchers last in their group, empty groups dropped. */
export function paletteGroups(tab: Tab, registry: CommandRegistry): PaletteGroup[] {
  const groups: PaletteGroup[] = []
  for (const group of tab.groups) {
    const rows: PaletteRow[] = []
    const seen = new Set<string>()
    const add = (control: Control) => {
      if (control.kind === "label" || seen.has(control.id) || !registry.isBuilt(control.id)) return
      seen.add(control.id)
      rows.push({
        id: control.id,
        label: "label" in control ? control.label : undefined,
        kind: control.kind === "check" ? "check" : control.kind === "field" ? "field" : "command",
      })
    }
    for (const column of group.columns) {
      for (const control of column.kind === "large" ? [column.control] : column.rows.flat()) add(control)
    }
    if (group.launcher && registry.isBuilt(group.launcher) && !seen.has(group.launcher))
      rows.push({ id: group.launcher, kind: "command" })
    if (rows.length) groups.push({ id: group.id, label: group.label, rows })
  }
  return groups
}

type Inside = { kind: "menu"; id: string } | { kind: "field"; id: string } | { kind: "tabs" } | null

export function CommandSheet({
  env,
  registry,
  tabs,
  tab: start,
  onClose,
}: {
  env: RibbonEnv
  registry: CommandRegistry
  tabs: Tab[]
  tab: string
  onClose: () => void
}) {
  // Pressed states (Bold, Wrap Text) follow the selection while the palette is open.
  useSyncExternalStore(env.c.subscribe, env.c.getVersion)
  const [tabId, setTabId] = useState(start)
  const [inside, setInside] = useState<Inside>(null)
  const tab = tabs.find((t) => t.id === tabId) ?? tabs[0]
  const groups = paletteGroups(tab, registry)
  const state: RibbonState = {
    env,
    registry,
    openMenu: null,
    setOpenMenu: () => undefined,
    keytips: null,
    afterRun: () => undefined,
  }

  const back = (title: string) => (
    <button type="button" className="phone-row is-back" onClick={() => setInside(null)}>
      <ChevronLeft size={20} aria-hidden />
      <span>{title}</span>
    </button>
  )

  let body: ReactNode
  if (inside?.kind === "tabs") {
    body = (
      <div role="listbox" aria-label="Tabs">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            role="option"
            aria-selected={t.id === tab.id}
            className="phone-row"
            onClick={() => {
              setTabId(t.id)
              setInside(null)
            }}
          >
            <span className="phone-row-icon">{t.id === tab.id ? <Check size={18} aria-hidden /> : null}</span>
            <span>{t.label}</span>
          </button>
        ))}
      </div>
    )
  } else if (inside?.kind === "menu") {
    const command = registry.get(inside.id)
    const entries = command?.menu ? (typeof command.menu === "function" ? command.menu(env) : command.menu) : []
    body = (
      <>
        {back(command?.labelFor?.(env) ?? command?.label ?? inside.id)}
        <MenuList entries={entries} close={() => setInside(null)} />
      </>
    )
  } else if (inside?.kind === "field") {
    const command = registry.get(inside.id)
    const field = command?.field
    const value = field?.value(env)
    body = (
      <>
        {back(command?.label ?? inside.id)}
        <div role="listbox" aria-label={command?.label}>
          {(field?.options?.(env) ?? []).map((option) => (
            <button
              key={option.value}
              type="button"
              role="option"
              aria-selected={option.value === value}
              className="phone-row"
              onClick={() => {
                field?.commit(env, option.value)
                setInside(null)
              }}
            >
              <span className="phone-row-icon">{option.value === value ? <Check size={18} aria-hidden /> : null}</span>
              <span className="phone-row-label">{option.preview ?? option.label}</span>
            </button>
          ))}
        </div>
      </>
    )
  } else {
    body = groups.map((group) => (
      <section key={group.id} className="phone-group" aria-label={group.label}>
        <h3 className="phone-group-title">{group.label}</h3>
        {group.rows.map((row) => (
          <PaletteRowView key={row.id} row={row} env={env} registry={registry} onInside={setInside} />
        ))}
      </section>
    ))
  }

  return (
    <BottomSheet
      label="Commands"
      onClose={onClose}
      header={
        <button
          type="button"
          className="phone-tab-picker"
          aria-haspopup="listbox"
          aria-expanded={inside?.kind === "tabs"}
          onClick={() => setInside(inside?.kind === "tabs" ? null : { kind: "tabs" })}
        >
          {tab.label}
          <ChevronDown size={16} aria-hidden />
        </button>
      }
    >
      <RibbonCtx.Provider value={state}>{body}</RibbonCtx.Provider>
    </BottomSheet>
  )
}

function PaletteRowView({
  row,
  env,
  registry,
  onInside,
}: {
  row: PaletteRow
  env: RibbonEnv
  registry: CommandRegistry
  onInside: (inside: Inside) => void
}) {
  const command = registry.get(row.id)!
  const enabled = registry.isEnabled(row.id, env)
  const label = (command.labelFor?.(env) ?? row.label ?? command.label).replace(/\.\.\.$/, "…")
  const icon = iconOf(command, env)
  const pressed = command.pressed?.(env)
  const entries = command.menu ? (typeof command.menu === "function" ? command.menu(env) : command.menu) : null

  if (row.kind === "check") {
    return (
      <button
        type="button"
        role="switch"
        aria-checked={!!pressed}
        className="phone-row"
        disabled={!enabled}
        onClick={() => command.run?.(env)}
      >
        <span className="phone-row-label">{label}</span>
        <span className="phone-switch" aria-hidden />
      </button>
    )
  }

  if (row.kind === "field" && command.field) {
    const field = command.field
    if (field.options) {
      return (
        <button
          type="button"
          className="phone-row"
          disabled={!enabled}
          onClick={() => onInside({ kind: "field", id: row.id })}
        >
          <span className="phone-row-label">{command.label}</span>
          <span className="phone-row-value">{field.value(env)}</span>
          <ChevronRight size={18} aria-hidden />
        </button>
      )
    }
    return (
      <label className="phone-row is-field">
        <span className="phone-row-label">{command.label}</span>
        <input
          className="phone-field"
          defaultValue={field.value(env)}
          disabled={!enabled}
          onKeyDown={(event) => event.key === "Enter" && event.currentTarget.blur()}
          onBlur={(event) => field.commit(env, event.currentTarget.value)}
        />
      </label>
    )
  }

  // A dropdown: the row opens its menu. A split button: the row runs, the › opens the menu.
  if (entries && !command.run) {
    return (
      <button
        type="button"
        className="phone-row"
        disabled={!enabled}
        onClick={() => onInside({ kind: "menu", id: row.id })}
      >
        <span className="phone-row-icon">{icon}</span>
        <span className="phone-row-label">{label}</span>
        <ChevronRight size={18} aria-hidden />
      </button>
    )
  }
  return (
    <div className="phone-row-split">
      <button
        type="button"
        className="phone-row"
        aria-pressed={pressed === undefined ? undefined : pressed}
        disabled={!enabled}
        onClick={() => command.run?.(env)}
      >
        <span className="phone-row-icon">{icon}</span>
        <span className="phone-row-label">{label}</span>
        {pressed ? <Check size={18} className="phone-row-check" aria-hidden /> : null}
      </button>
      {entries ? (
        <button
          type="button"
          className="phone-row-more"
          aria-label={`${label} options`}
          disabled={!enabled}
          onClick={() => onInside({ kind: "menu", id: row.id })}
        >
          <ChevronRight size={18} aria-hidden />
        </button>
      ) : null}
    </div>
  )
}
