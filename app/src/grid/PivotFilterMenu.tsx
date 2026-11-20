import { wildcardRegExp, type PivotChoices, type PivotFilterTarget } from "@cells/core"
import { ArrowDownAZ, ArrowUpAZ, Check, ChevronRight, FunnelX, Search } from "lucide-react"
import { useEffect, useLayoutEffect, useRef, useState } from "react"
import { createPortal } from "react-dom"
import type { SheetController } from "@/sheet/controller"
import type { PivotFilterRequest } from "@/sheet/pivots"
import { LABEL_OPS, VALUE_OPS } from "@/shell/PivotFilterDialogs"
import { cn } from "@/lib/cn"
import { Submenu, type SubItem } from "./FilterMenu"

/**
 * The dropdown of a PivotTable's filter button, as Excel's. On Row Labels / Column Labels or a field's header: Select
 * field (when the area has several), the sorts, Clear Filter, a search box and the items to tick. On a report filter:
 * the search box, (All) and the items to pick one of, or to tick with Select Multiple Items.
 */
export function PivotFilterMenu({
  controller,
  row,
  col,
  target,
  left,
  top,
  sheet,
  onClose,
}: {
  controller: SheetController
  row: number
  col: number
  /** The PivotTable's sheet when a PivotChart's button opened the dropdown (it may not be the sheet shown). */
  sheet?: number
  target: PivotFilterTarget
  left: number
  top: number
  onClose: () => void
}) {
  const ref = useRef<HTMLDivElement>(null)
  const pivots = controller.pivots.on(sheet)
  const page = target.kind === "page"
  const [field, setField] = useState(target.fields[0])
  const [choices, setChoices] = useState<PivotChoices | null>(() => pivots.choices(row, col, field))
  const [ticked, setTicked] = useState<Set<number>>(() => tickedOf(choices))
  /** Report filters: the one item picked (null: (All)). */
  const [pick, setPick] = useState<number | null>(choices?.picked ?? null)
  const [multiple, setMultiple] = useState(Boolean(choices?.multipleItems && choices.picked === null))
  const [search, setSearch] = useState("")
  const [place, setPlace] = useState({ left, top, visible: false })
  const [submenu, setSubmenu] = useState<"label" | "value" | null>(null)

  useLayoutEffect(() => {
    const menu = ref.current
    if (!menu) return
    const rect = menu.getBoundingClientRect()
    const x = Math.max(4, Math.min(left, window.innerWidth - rect.width - 4))
    const y = top + rect.height > window.innerHeight - 4 ? Math.max(4, window.innerHeight - rect.height - 4) : top
    setPlace({ left: x, top: y, visible: true })
    menu.querySelector<HTMLElement>(".fm-row:not(:disabled), select, input")?.focus({ preventScroll: true })
  }, [left, top])

  useEffect(() => {
    const onDown = (event: PointerEvent) => {
      const target = event.target as HTMLElement
      if (!ref.current?.contains(target) && !target.closest?.(".fm-submenu")) onClose()
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault()
        event.stopPropagation()
        if (submenu) setSubmenu(null)
        else onClose()
      }
    }
    window.addEventListener("pointerdown", onDown, true)
    window.addEventListener("keydown", onKey, true)
    window.addEventListener("blur", onClose)
    return () => {
      window.removeEventListener("pointerdown", onDown, true)
      window.removeEventListener("keydown", onKey, true)
      window.removeEventListener("blur", onClose)
    }
  }, [onClose, submenu])

  if (!choices) return null
  const changeField = (f: number) => {
    const next = pivots.choices(row, col, f)
    setField(f)
    setChoices(next)
    setTicked(tickedOf(next))
    setSearch("")
  }

  const matcher = search.trim() ? wildcardRegExp(`*${search.trim()}*`) : null
  const shown = matcher ? choices.items.filter((i) => matcher.test(i.label.toLowerCase())) : choices.items
  const checkboxes = !page || multiple
  const every = shown.length > 0 && shown.every((i) => ticked.has(i.x))
  const some = shown.some((i) => ticked.has(i.x))

  const ok = () => {
    onClose()
    if (checkboxes) {
      // While searching, only the matches that are ticked count, as in Excel.
      const keep = new Set(matcher ? shown.filter((i) => ticked.has(i.x)).map((i) => i.x) : ticked)
      if (!keep.size) return
      const all = choices.items.every((i) => keep.has(i.x))
      pivots.showItems(row, col, field, all && !page ? null : keep)
    } else pivots.pickPage(row, col, field, pick)
  }
  const sort = (descending: boolean) => {
    onClose()
    pivots.sortField(row, col, field, descending ? "descending" : "ascending")
  }
  const sortLabels = {
    text: ["Sort A to Z", "Sort Z to A"],
    number: ["Sort Smallest to Largest", "Sort Largest to Smallest"],
    date: ["Sort Oldest to Newest", "Sort Newest to Oldest"],
  }[choices.kind]

  // ── Label Filters and Value Filters ──
  const dialog = (kind: PivotFilterRequest["kind"], op?: string) => {
    onClose()
    controller.shell.pivotFilterDialog?.({ kind, row, col, field, op, sheet })
  }
  const current = choices.filter
  /** A label filter belongs to Label Filters; a value filter and Top 10 to Value Filters. */
  const isOf = (kind: "label" | "value") => Boolean(current) && (current!.kind === "label") === (kind === "label")
  const subItems = (kind: "label" | "value"): SubItem[] => {
    const out: SubItem[] = []
    if (isOf(kind))
      out.push(
        {
          kind: "item",
          label: "Clear Filter",
          run: () => {
            onClose()
            pivots.setFilter(row, col, field, null)
          },
        },
        { kind: "separator" },
      )
    for (const o of kind === "label" ? LABEL_OPS : VALUE_OPS)
      out.push({
        kind: "item",
        label: o.menu,
        checked: current?.kind === kind && current.op === o.op,
        run: () => dialog(kind, o.op),
      })
    if (kind === "value")
      out.push(
        { kind: "separator" },
        { kind: "item", label: "Top 10...", checked: current?.kind === "top", run: () => dialog("top") },
      )
    return out
  }
  const hasValues = pivots.valueNames(row, col).length > 0
  const subRow = (id: "label" | "value", label: string, disabled = false) => (
    <button
      type="button"
      role="menuitem"
      className={cn("fm-row", submenu === id && "is-open")}
      disabled={disabled}
      aria-haspopup="menu"
      aria-expanded={submenu === id}
      data-sub={id}
      onPointerEnter={() => setSubmenu(id)}
      onKeyDown={(event) => {
        if (event.key === "ArrowRight" || event.key === "Enter") {
          event.preventDefault()
          setSubmenu(id)
        }
      }}
      onClick={() => setSubmenu(id)}
    >
      <span className="fm-row-icon">{isOf(id) ? <Check size={14} /> : null}</span>
      <span>{label}</span>
      <ChevronRight className="fm-row-more" size={14} strokeWidth={1.6} />
    </button>
  )

  const okDisabled = checkboxes ? !(matcher ? shown.some((i) => ticked.has(i.x)) : ticked.size) : false
  const tick = (x: number, on: boolean) => {
    const next = new Set(ticked)
    if (on) next.add(x)
    else next.delete(x)
    setTicked(next)
  }

  return createPortal(
    <div
      ref={ref}
      className="filter-menu pivot-filter-menu"
      role="menu"
      aria-label={`Filter ${choices.name}`}
      style={{ left: place.left, top: place.top, visibility: place.visible ? "visible" : "hidden" }}
      onContextMenu={(event) => event.preventDefault()}
      onKeyDown={(event) => {
        const rows = [...(ref.current?.querySelectorAll<HTMLButtonElement>(".fm-row:not(:disabled)") ?? [])]
        const at = rows.indexOf(document.activeElement as HTMLButtonElement)
        if (at < 0) return
        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
          event.preventDefault()
          rows[(at + (event.key === "ArrowDown" ? 1 : -1) + rows.length) % rows.length]?.focus()
        }
      }}
    >
      {!page && target.fields.length > 1 ? (
        <label className="pfm-field">
          <span>Select field:</span>
          <select value={field} onChange={(event) => changeField(Number(event.target.value))}>
            {target.fields.map((f) => (
              <option key={f} value={f}>
                {pivots.choices(row, col, f)?.name ?? ""}
              </option>
            ))}
          </select>
        </label>
      ) : null}
      {!page ? (
        <>
          <button
            type="button"
            role="menuitemradio"
            aria-checked={choices.sort === "ascending"}
            className="fm-row"
            onClick={() => sort(false)}
          >
            <span className="fm-row-icon">
              <ArrowUpAZ size={15} strokeWidth={1.5} />
            </span>
            <span>{sortLabels[0]}</span>
          </button>
          <button
            type="button"
            role="menuitemradio"
            aria-checked={choices.sort === "descending"}
            className="fm-row"
            onClick={() => sort(true)}
          >
            <span className="fm-row-icon">
              <ArrowDownAZ size={15} strokeWidth={1.5} />
            </span>
            <span>{sortLabels[1]}</span>
          </button>
          <div className="menu-separator" role="separator" />
          <button
            type="button"
            role="menuitem"
            className="fm-row"
            disabled={!choices.filtered}
            onPointerEnter={() => setSubmenu(null)}
            onClick={() => {
              onClose()
              pivots.clearField(row, col, field)
            }}
          >
            <span className="fm-row-icon">
              <FunnelX size={15} strokeWidth={1.5} />
            </span>
            <span>{`Clear Filter From "${choices.name}"`}</span>
          </button>
          {subRow("label", "Label Filters")}
          {subRow("value", "Value Filters", !hasValues)}
        </>
      ) : null}
      <label className="fm-search" onPointerEnter={() => setSubmenu(null)}>
        <input
          type="search"
          value={search}
          placeholder="Search"
          aria-label="Search"
          onChange={(event) => {
            setSearch(event.target.value)
            // Searching ticks every match, as Excel's list does.
            const m = event.target.value.trim() ? wildcardRegExp(`*${event.target.value.trim()}*`) : null
            if (m && checkboxes)
              setTicked(new Set(choices.items.filter((i) => m.test(i.label.toLowerCase())).map((i) => i.x)))
            else if (!m) setTicked(tickedOf(choices))
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault()
              if (!okDisabled) ok()
            }
          }}
        />
        <Search size={14} strokeWidth={1.6} aria-hidden />
      </label>
      <div className="fm-list" role={checkboxes ? "group" : "radiogroup"} aria-label="Items">
        {!shown.length ? (
          <p className="fm-empty">No matches</p>
        ) : checkboxes ? (
          <>
            <TriBox
              checked={every}
              mixed={some && !every}
              label={matcher ? "(Select All Search Results)" : "(Select All)"}
              onChange={(on) => {
                const next = new Set(ticked)
                for (const i of shown)
                  if (on) next.add(i.x)
                  else next.delete(i.x)
                setTicked(next)
              }}
            />
            {shown.map((i) => (
              <TriBox key={i.x} checked={ticked.has(i.x)} label={i.label} onChange={(on) => tick(i.x, on)} />
            ))}
          </>
        ) : (
          <>
            {!matcher ? <PickRow label="(All)" picked={pick === null} onPick={() => setPick(null)} onOk={ok} /> : null}
            {shown.map((i) => (
              <PickRow key={i.x} label={i.label} picked={pick === i.x} onPick={() => setPick(i.x)} onOk={ok} />
            ))}
          </>
        )}
      </div>
      {page ? (
        <label className="fm-check pfm-multiple">
          <input
            type="checkbox"
            checked={multiple}
            onChange={(event) => {
              setMultiple(event.target.checked)
              // Switching to tick boxes starts from the pick (all of them for (All)).
              if (event.target.checked) setTicked(new Set(pick === null ? choices.items.map((i) => i.x) : [pick]))
            }}
          />
          <span>Select Multiple Items</span>
        </label>
      ) : null}
      <div className="fm-actions">
        <button type="button" className="dialog-btn is-primary" disabled={okDisabled} onClick={ok}>
          OK
        </button>
        <button type="button" className="dialog-btn" onClick={onClose}>
          Cancel
        </button>
      </div>
      {submenu ? (
        <Submenu
          parent={ref.current?.querySelector<HTMLElement>(`[data-sub="${submenu}"]`) ?? null}
          items={subItems(submenu)}
          onBack={() => {
            ref.current?.querySelector<HTMLElement>(`[data-sub="${submenu}"]`)?.focus()
            setSubmenu(null)
          }}
        />
      ) : null}
    </div>,
    document.body,
  )
}

/** The items ticked as the menu opens (a report filter's single pick: just it). */
function tickedOf(choices: PivotChoices | null): Set<number> {
  return new Set((choices?.items ?? []).filter((i) => i.checked).map((i) => i.x))
}

function TriBox({
  checked,
  mixed,
  label,
  onChange,
}: {
  checked: boolean
  mixed?: boolean
  label: string
  onChange: (v: boolean) => void
}) {
  const ref = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = Boolean(mixed)
  }, [mixed])
  return (
    <label className="fm-check">
      <input ref={ref} type="checkbox" checked={checked && !mixed} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

/** One line of a report filter's single-pick list: a click picks it, a double-click picks it and closes. */
function PickRow({
  label,
  picked,
  onPick,
  onOk,
}: {
  label: string
  picked: boolean
  onPick: () => void
  onOk: () => void
}) {
  return (
    <button
      type="button"
      role="radio"
      aria-checked={picked}
      className={`pfm-pick${picked ? " is-picked" : ""}`}
      onClick={onPick}
      onDoubleClick={onOk}
    >
      {label}
    </button>
  )
}
