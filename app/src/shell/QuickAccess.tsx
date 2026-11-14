import { Check, ChevronDown, Search } from "lucide-react"
import { useEffect, useRef, useState } from "react"
import { QAT_CHOICES } from "@/ribbon/qat-commands"
import { iconOf, NOT_BUILT, shortcutText, type CommandRegistry, type RibbonEnv } from "@/ribbon/registry"
import { searchCommands } from "@/ribbon/tell-me"
import { ContextMenu } from "./ContextMenu"

type QatProps = {
  registry: CommandRegistry
  env: RibbonEnv
  ids: string[]
  onChange: (ids: string[]) => void
  /** Commands are run from the title bar: give the sheet its focus back after. */
  onDone: () => void
}

/**
 * The Quick Access Toolbar: the user's commands (Undo and Redo to start with) and Excel's Customize menu, which
 * ticks commands on and off. Right-click a button to remove it; right-click a ribbon command to add it.
 */
export function QuickAccess({ registry, env, ids, onChange, onDone }: QatProps) {
  const [menu, setMenu] = useState<{ x: number; y: number; id?: string } | null>(null)
  const toggle = (id: string) => onChange(ids.includes(id) ? ids.filter((i) => i !== id) : [...ids, id])
  return (
    <div className="qat" role="toolbar" aria-label="Quick Access Toolbar">
      {ids.map((id, i) => {
        const command = registry.get(id)
        if (!command) return null
        const enabled = registry.isEnabled(id, env)
        const label = command.labelFor?.(env) ?? command.label
        const shortcut = shortcutText(command.shortcut)
        return (
          <button
            type="button"
            key={id}
            className="titlebar-button"
            aria-label={label}
            title={
              registry.isBuilt(id)
                ? `${label}${shortcut ? ` (${shortcut})` : ""} — Alt, ${i + 1}`
                : `${label} — ${NOT_BUILT}`
            }
            disabled={!enabled}
            data-qat={id}
            onClick={() => {
              command.run?.(env)
              onDone()
            }}
            onContextMenu={(event) => {
              event.preventDefault()
              setMenu({ x: event.clientX, y: event.clientY, id })
            }}
          >
            {iconOf(command, env) ?? <span className="qat-text">{label.slice(0, 2)}</span>}
          </button>
        )
      })}
      <button
        type="button"
        className="titlebar-button qat-customize"
        aria-label="Customize Quick Access Toolbar"
        title="Customize Quick Access Toolbar"
        aria-haspopup="menu"
        onClick={(event) => {
          const r = event.currentTarget.getBoundingClientRect()
          setMenu({ x: r.left, y: r.bottom + 2 })
        }}
      >
        <ChevronDown size={13} strokeWidth={1.75} />
      </button>
      {menu ? (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={
            menu.id
              ? [
                  {
                    kind: "item",
                    label: "Remove from Quick Access Toolbar",
                    onSelect: () => onChange(ids.filter((i) => i !== menu.id)),
                  },
                ]
              : [
                  ...QAT_CHOICES.map((id) => ({
                    kind: "item" as const,
                    label: registry.get(id)?.label ?? id,
                    icon: ids.includes(id) ? <Check size={14} strokeWidth={2} /> : undefined,
                    onSelect: () => toggle(id),
                  })),
                  { kind: "separator" as const },
                  {
                    kind: "item" as const,
                    label: "Reset to Default",
                    onSelect: () => onChange(["Undo", "Redo"]),
                  },
                ]
          }
        />
      ) : null}
    </div>
  )
}

/**
 * Search (Alt+Q), Excel's Tell Me: type what you want to do; the matching commands run from the list (Enter runs
 * the highlighted one). A command that opens a menu or a box shows it on its ribbon tab.
 */
export function TellMe({
  registry,
  env,
  onShow,
  onDone,
}: {
  registry: CommandRegistry
  env: RibbonEnv
  onShow: (id: string) => boolean
  onDone: () => void
}) {
  const [query, setQuery] = useState("")
  const [pick, setPick] = useState(0)
  const [open, setOpen] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const results = open ? searchCommands(registry, env, query) : []

  // Alt+Q from anywhere.
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.altKey && !event.ctrlKey && !event.metaKey && event.key.toLowerCase() === "q") {
        event.preventDefault()
        event.stopPropagation()
        inputRef.current?.focus()
        inputRef.current?.select()
      }
    }
    window.addEventListener("keydown", onKey, true)
    return () => window.removeEventListener("keydown", onKey, true)
  }, [])

  const leave = () => {
    setQuery("")
    setOpen(false)
    setPick(0)
  }
  const choose = (id: string) => {
    const command = registry.get(id)
    if (!command || !registry.isEnabled(id, env)) return
    leave()
    if (command.run) {
      command.run(env)
      onDone()
    } else if (!onShow(id)) onDone()
  }

  return (
    <div className="tell-me">
      <Search size={14} strokeWidth={1.75} className="tell-me-icon" aria-hidden />
      <input
        ref={inputRef}
        className="tell-me-input"
        type="search"
        placeholder="Search (Alt+Q)"
        aria-label="Search"
        aria-autocomplete="list"
        aria-expanded={results.length > 0}
        aria-controls="tell-me-results"
        value={query}
        onFocus={() => setOpen(true)}
        onBlur={() => setOpen(false)}
        onChange={(event) => {
          setQuery(event.target.value)
          setPick(0)
          setOpen(true)
        }}
        onKeyDown={(event) => {
          event.stopPropagation()
          if (event.key === "ArrowDown" || event.key === "ArrowUp") {
            event.preventDefault()
            if (results.length)
              setPick((p) => (p + (event.key === "ArrowDown" ? 1 : -1) + results.length) % results.length)
          } else if (event.key === "Enter") {
            event.preventDefault()
            const hit = results[pick]
            if (hit) choose(hit.id)
          } else if (event.key === "Escape") {
            event.preventDefault()
            leave()
            onDone()
          }
        }}
      />
      {results.length ? (
        <div className="tell-me-results" id="tell-me-results" role="listbox" aria-label="Search results">
          <div className="tell-me-heading">Actions</div>
          {results.map((command, i) => {
            const built = registry.isBuilt(command.id)
            const enabled = registry.isEnabled(command.id, env)
            const shortcut = shortcutText(command.shortcut)
            return (
              <button
                type="button"
                role="option"
                key={command.id}
                aria-selected={i === pick}
                aria-disabled={!enabled || undefined}
                className={`tell-me-row${i === pick ? " is-picked" : ""}`}
                title={built ? command.description : NOT_BUILT}
                onPointerDown={(event) => event.preventDefault()}
                onPointerEnter={() => setPick(i)}
                onClick={() => choose(command.id)}
              >
                <span className="menu-row-icon">{iconOf(command, env)}</span>
                <span className="tell-me-label">{command.labelFor?.(env) ?? command.label}</span>
                {shortcut ? <span className="menu-row-shortcut">{shortcut}</span> : null}
              </button>
            )
          })}
        </div>
      ) : null}
    </div>
  )
}
