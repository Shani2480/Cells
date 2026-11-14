import { Check, ChevronRight } from "lucide-react"
import { Fragment, useState } from "react"
import { Popover } from "@/ui/popover"
import { useRibbon } from "./context"
import { NOT_BUILT, iconOf, shortcutText, type MenuEntry } from "./registry"

/**
 * A ribbon dropdown's contents: command rows (picture, name, shortcut, a tick when on), headings, separators and
 * custom panels. A row whose command has its own menu opens it to the side.
 */
export function MenuList({ entries, close }: { entries: MenuEntry[]; close: () => void }) {
  const { env, registry, afterRun } = useRibbon()
  return (
    <div className="ribbon-menu" role="menu">
      {entries.map((entry, i) => {
        if (entry === "-") return <div key={i} className="menu-separator" role="separator" />
        if (typeof entry === "object" && "heading" in entry)
          return (
            <div key={i} className="menu-heading">
              {entry.heading}
            </div>
          )
        if (typeof entry === "object") return <Fragment key={i}>{entry.panel(env, close)}</Fragment>
        const command = registry.get(entry)
        if (!command) return null
        const built = registry.isBuilt(entry)
        const enabled = registry.isEnabled(entry, env)
        if (command.menu && !command.run) {
          return <SubmenuRow key={entry} id={entry} closeAll={close} />
        }
        const pressed = command.pressed?.(env)
        return (
          <button
            key={entry}
            type="button"
            role={pressed === undefined ? "menuitem" : "menuitemcheckbox"}
            aria-checked={pressed}
            className="menu-row"
            disabled={!enabled}
            title={built ? undefined : NOT_BUILT}
            onClick={() => {
              close()
              command.run?.(env)
              afterRun()
            }}
          >
            <span className="menu-row-icon">
              {iconOf(command, env) ?? (pressed ? <Check size={14} strokeWidth={2} /> : null)}
            </span>
            <span>{command.labelFor?.(env) ?? command.label}</span>
            {command.shortcut ? <span className="menu-row-shortcut">{shortcutText(command.shortcut)}</span> : null}
          </button>
        )
      })}
    </div>
  )
}

function SubmenuRow({ id, closeAll }: { id: string; closeAll: () => void }) {
  const { env, registry } = useRibbon()
  const [open, setOpen] = useState(false)
  const command = registry.get(id)!
  const entries = typeof command.menu === "function" ? command.menu(env) : command.menu!
  return (
    <Popover
      open={open}
      onOpenChange={setOpen}
      side="right"
      trigger={
        <button
          type="button"
          role="menuitem"
          aria-haspopup="menu"
          className="menu-row"
          onMouseEnter={() => setOpen(true)}
        >
          <span className="menu-row-icon">{iconOf(command, env)}</span>
          <span>{command.label}</span>
          <ChevronRight className="menu-row-more" size={13} strokeWidth={1.75} />
        </button>
      }
    >
      <MenuList
        entries={entries}
        close={() => {
          setOpen(false)
          closeAll()
        }}
      />
    </Popover>
  )
}
