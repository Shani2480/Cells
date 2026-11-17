import {
  displayText,
  filterChoices,
  sameIcon as sameIconRef,
  valuesFilter,
  wildcardRegExp,
  type ColumnFilter,
  type DynamicFilter,
  type FilterChoices,
  type IconRef,
} from "@cells/core"
import { CfIcon } from "./CfIcon"
import { ArrowDownAZ, ArrowUpAZ, Check, ChevronRight, FunnelX, Minus, Plus, Search } from "lucide-react"
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react"
import { createPortal } from "react-dom"
import type { CustomPreset, FilterDialogRequest, SortBy } from "@/sheet/data"
import type { SheetController } from "@/sheet/controller"
import { todaySerial } from "@/sheet/controller"
import { cn } from "@/lib/cn"

const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
]

type SubItem =
  | { kind: "separator" }
  | {
      kind: "item"
      label: string
      checked?: boolean
      run?: () => void
      sub?: SubItem[]
      swatch?: string | null
      /** A conditional formatting icon, or null for No Cell Icon. */
      icon?: IconRef | null
    }

/**
 * A filter button's dropdown, as Excel's: sorts, Clear Filter, Filter by Color, the Text / Number / Date Filters
 * submenu, a search box and the list of values to tick, then OK and Cancel.
 */
export function FilterMenu({
  controller,
  col,
  left,
  top,
  onClose,
  onDialog,
}: {
  controller: SheetController
  col: number
  left: number
  top: number
  onClose: () => void
  onDialog: (request: FilterDialogRequest) => void
}) {
  const ref = useRef<HTMLDivElement>(null)
  const filter = controller.data.autoFilter!
  const offset = col - filter.range.c0
  const current = filter.columns[offset]
  const sheet = controller.sheet
  const styles = controller.wb.styles
  // The menu reads the sheet once, when it opens.
  const [initial] = useState(() => filterChoices(sheet, styles, filter, col, todaySerial(), controller.data.look))
  const [choices, setChoices] = useState<FilterChoices>(initial)
  const [search, setSearch] = useState("")
  const [addToFilter, setAddToFilter] = useState(false)
  /** While searching: the matching entries, and those still ticked. */
  const [found, setFound] = useState<string[] | null>(null)
  const [ticked, setTicked] = useState<Set<string>>(new Set())
  const [submenu, setSubmenu] = useState<string | null>(null)
  const [place, setPlace] = useState({ left, top, visible: false })

  const headerText = (() => {
    const data = sheet.cells.get(filter.range.r0, col)
    const code = styles.get(sheet.styleAt(filter.range.r0, col)).numberFormat
    return displayText(data?.value, code)
  })()

  useLayoutEffect(() => {
    const menu = ref.current
    if (!menu) return
    const rect = menu.getBoundingClientRect()
    const x = Math.max(4, Math.min(left, window.innerWidth - rect.width - 4))
    const y = top + rect.height > window.innerHeight - 4 ? Math.max(4, window.innerHeight - rect.height - 4) : top
    setPlace({ left: x, top: y, visible: true })
    menu.querySelector<HTMLButtonElement>(".fm-row:not(:disabled)")?.focus({ preventScroll: true })
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

  const apply = (column: ColumnFilter | null) => {
    onClose()
    controller.data.filterColumn(col, column)
  }
  const sort = (descending: boolean, by?: SortBy) => {
    onClose()
    controller.data.sortFiltered(col, descending, by)
  }
  const dialog = (request: FilterDialogRequest) => {
    onClose()
    onDialog(request)
  }

  // ── Search: the list narrows to matches, all ticked; OK filters to them (plus the current ones, if asked). ──
  const pattern = search.trim()
  const matcher = pattern ? wildcardRegExp(`*${pattern}*`) : null
  const runSearch = (text: string) => {
    setSearch(text)
    const p = text.trim()
    if (!p) {
      setFound(null)
      return
    }
    const m = wildcardRegExp(`*${p}*`)
    const hits: string[] = []
    for (const d of initial.dates) if (m.test(dateText(d).toLowerCase())) hits.push(dateText(d))
    for (const v of initial.values) if (m.test(v.text.trim().toLowerCase())) hits.push(v.text)
    setFound(hits)
    setTicked(new Set(hits))
  }

  const ok = () => {
    if (matcher && found) {
      const keep = (checked: boolean, text: string) => ticked.has(text) || (addToFilter && checked)
      const next: FilterChoices = {
        ...choices,
        values: initial.values.map((v) => ({ ...v, checked: keep(v.checked && Boolean(current), v.text) })),
        dates: initial.dates.map((d) => ({ ...d, checked: keep(d.checked && Boolean(current), dateText(d)) })),
        blanks: initial.blanks ? { checked: addToFilter && Boolean(current) && initial.blanks.checked } : null,
      }
      if (!ticked.size && !addToFilter) return onClose()
      return apply(valuesFilter(next))
    }
    apply(valuesFilter(choices))
  }

  const anyChecked =
    choices.values.some((v) => v.checked) || choices.dates.some((d) => d.checked) || Boolean(choices.blanks?.checked)

  // ── Submenus ──
  const isDynamic = (type: DynamicFilter) => current?.kind === "dynamic" && current.type === type
  const dyn = (label: string, type: DynamicFilter): SubItem => ({
    kind: "item",
    label,
    checked: isDynamic(type),
    run: () => apply({ kind: "dynamic", type }),
  })
  const custom = (label: string, preset: CustomPreset): SubItem => ({
    kind: "item",
    label,
    checked: false,
    run: () => dialog({ kind: "custom", col, preset }),
  })
  const sep: SubItem = { kind: "separator" }
  const customFilter: SubItem = {
    kind: "item",
    label: "Custom Filter...",
    checked: current?.kind === "custom",
    run: () => dialog({ kind: "custom", col }),
  }
  const filters: Record<FilterChoices["kind"], SubItem[]> = {
    text: [
      custom("Equals...", { first: "equals" }),
      custom("Does Not Equal...", { first: "notEquals" }),
      sep,
      custom("Begins With...", { first: "beginsWith" }),
      custom("Ends With...", { first: "endsWith" }),
      sep,
      custom("Contains...", { first: "contains" }),
      custom("Does Not Contain...", { first: "notContains" }),
      sep,
      customFilter,
    ],
    number: [
      custom("Equals...", { first: "equals" }),
      custom("Does Not Equal...", { first: "notEquals" }),
      sep,
      custom("Greater Than...", { first: "greater" }),
      custom("Greater Than Or Equal To...", { first: "greaterOrEqual" }),
      custom("Less Than...", { first: "less" }),
      custom("Less Than Or Equal To...", { first: "lessOrEqual" }),
      custom("Between...", { first: "greaterOrEqual", second: "lessOrEqual", and: true }),
      sep,
      {
        kind: "item",
        label: "Top 10...",
        checked: current?.kind === "top10",
        run: () => dialog({ kind: "top10", col }),
      },
      dyn("Above Average", "aboveAverage"),
      dyn("Below Average", "belowAverage"),
      sep,
      customFilter,
    ],
    date: [
      custom("Equals...", { first: "equals" }),
      sep,
      custom("Before...", { first: "less" }),
      custom("After...", { first: "greater" }),
      custom("Between...", { first: "greaterOrEqual", second: "lessOrEqual", and: true }),
      sep,
      dyn("Tomorrow", "tomorrow"),
      dyn("Today", "today"),
      dyn("Yesterday", "yesterday"),
      sep,
      dyn("Next Week", "nextWeek"),
      dyn("This Week", "thisWeek"),
      dyn("Last Week", "lastWeek"),
      sep,
      dyn("Next Month", "nextMonth"),
      dyn("This Month", "thisMonth"),
      dyn("Last Month", "lastMonth"),
      sep,
      dyn("Next Quarter", "nextQuarter"),
      dyn("This Quarter", "thisQuarter"),
      dyn("Last Quarter", "lastQuarter"),
      sep,
      dyn("Next Year", "nextYear"),
      dyn("This Year", "thisYear"),
      dyn("Last Year", "lastYear"),
      sep,
      dyn("Year to Date", "yearToDate"),
      {
        kind: "item",
        label: "All Dates in the Period",
        sub: [
          ...[1, 2, 3, 4].map((q) => dyn(`Quarter ${q}`, `Q${q}` as DynamicFilter)),
          ...MONTHS.map((m, i) => dyn(m, `M${i + 1}` as DynamicFilter)),
        ],
      },
      sep,
      customFilter,
    ],
  }
  const filtersLabel = { text: "Text Filters", number: "Number Filters", date: "Date Filters" }[choices.kind]
  const sortLabels = {
    text: ["Sort A to Z", "Sort Z to A"],
    number: ["Sort Smallest to Largest", "Sort Largest to Smallest"],
    date: ["Sort Oldest to Newest", "Sort Newest to Oldest"],
  }[choices.kind]
  const colorItems = (on: "cell" | "font", run: (color: string | null) => void): SubItem[] => {
    const list = on === "cell" ? initial.cellColors : initial.fontColors
    return list.map((color) => ({
      kind: "item",
      label: color === null ? (on === "cell" ? "No Fill" : "Automatic") : "",
      swatch: on === "cell" ? (color ?? "transparent") : color,
      checked:
        current?.kind === "color" && current.cell === (on === "cell") && (current.color ?? null) === (color ?? null),
      run: () => run(color),
    }))
  }
  const colorMenu = (byFilter: boolean): SubItem[] => {
    const make = (on: "cell" | "font") =>
      colorItems(on, (color) =>
        byFilter ? apply({ kind: "color", cell: on === "cell", color }) : sort(false, { on, color }),
      )
    const cells = initial.cellColors.length > 1 ? make("cell") : []
    const fonts = initial.fontColors.length > 1 ? make("font") : []
    const out: SubItem[] = []
    if (cells.length) out.push({ kind: "item", label: `${byFilter ? "Filter" : "Sort"} by Cell Color` }, ...cells)
    if (fonts.length) {
      if (out.length) out.push(sep)
      out.push({ kind: "item", label: `${byFilter ? "Filter" : "Sort"} by Font Color` }, ...fonts)
    }
    if (initial.icons.length) {
      if (out.length) out.push(sep)
      out.push(
        { kind: "item", label: `${byFilter ? "Filter" : "Sort"} by Cell Icon` },
        ...initial.icons.map((icon): SubItem => ({
          kind: "item",
          label: icon ? "" : "No Cell Icon",
          icon,
          checked: current?.kind === "icon" && sameIconRef(current.icon, icon),
          run: () => (byFilter ? apply({ kind: "icon", icon }) : sort(false, { on: "icon", icon })),
        })),
      )
    }
    return out
  }
  const hasColors = initial.cellColors.length > 1 || initial.fontColors.length > 1 || initial.icons.length > 0

  const rowButton = (
    id: string,
    label: string,
    options: { icon?: ReactNode; disabled?: boolean; sub?: boolean; run?: () => void },
  ) => (
    <button
      type="button"
      role="menuitem"
      className={cn("fm-row", submenu === id && "is-open")}
      disabled={options.disabled}
      aria-haspopup={options.sub ? "menu" : undefined}
      aria-expanded={options.sub ? submenu === id : undefined}
      data-sub={options.sub ? id : undefined}
      onPointerEnter={() => setSubmenu(options.sub ? id : null)}
      onKeyDown={(event) => {
        if (options.sub && (event.key === "ArrowRight" || event.key === "Enter")) {
          event.preventDefault()
          setSubmenu(id)
        }
      }}
      onClick={() => (options.sub ? setSubmenu(id) : options.run?.())}
    >
      <span className="fm-row-icon">{options.icon}</span>
      <span>{label}</span>
      {options.sub ? <ChevronRight className="fm-row-more" size={14} strokeWidth={1.6} /> : null}
    </button>
  )

  const submenuItems =
    submenu === "filters"
      ? filters[choices.kind]
      : submenu === "sortColor"
        ? colorMenu(false)
        : submenu === "filterColor"
          ? colorMenu(true)
          : null

  return createPortal(
    <div
      ref={ref}
      className="filter-menu"
      role="menu"
      aria-label={`Filter ${headerText}`}
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
      {rowButton("asc", sortLabels[0], {
        icon: <ArrowUpAZ size={15} strokeWidth={1.5} />,
        run: () => sort(false),
      })}
      {rowButton("desc", sortLabels[1], {
        icon: <ArrowDownAZ size={15} strokeWidth={1.5} />,
        run: () => sort(true),
      })}
      {rowButton("sortColor", "Sort by Color", { sub: true, disabled: !hasColors })}
      <div className="menu-separator" role="separator" />
      {rowButton("clear", `Clear Filter From "${headerText}"`, {
        icon: <FunnelX size={15} strokeWidth={1.5} />,
        disabled: !current,
        run: () => apply(null),
      })}
      {rowButton("filterColor", "Filter by Color", { sub: true, disabled: !hasColors })}
      {rowButton("filters", filtersLabel, {
        sub: true,
        icon: current && current.kind !== "values" ? <Check size={14} /> : null,
      })}
      <label className="fm-search">
        <input
          type="search"
          value={search}
          placeholder="Search"
          aria-label="Search"
          onPointerEnter={() => setSubmenu(null)}
          onChange={(event) => runSearch(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault()
              ok()
            }
          }}
        />
        <Search size={14} strokeWidth={1.6} aria-hidden />
      </label>
      <div className="fm-list" role="tree" aria-label="Values" onPointerEnter={() => setSubmenu(null)}>
        {found ? (
          <SearchList
            hits={found}
            ticked={ticked}
            onTick={setTicked}
            addToFilter={addToFilter}
            onAddToFilter={setAddToFilter}
          />
        ) : (
          <ValueTree choices={choices} onChange={setChoices} />
        )}
      </div>
      <div className="fm-actions">
        <button
          type="button"
          className="dialog-btn is-primary"
          disabled={found ? !ticked.size && !addToFilter : !anyChecked}
          onClick={ok}
        >
          OK
        </button>
        <button type="button" className="dialog-btn" onClick={onClose}>
          Cancel
        </button>
      </div>
      {submenuItems && submenu ? (
        <Submenu
          parent={ref.current?.querySelector<HTMLElement>(`[data-sub="${submenu}"]`) ?? null}
          items={submenuItems}
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

const dateText = (d: { year: number; month: number; day: number }) => `${MONTHS[d.month - 1]} ${d.day}, ${d.year}`

/** A flyout beside its row; nested flyouts (All Dates in the Period) open the same way. */
function Submenu({ parent, items, onBack }: { parent: HTMLElement | null; items: SubItem[]; onBack: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const [place, setPlace] = useState<{ left: number; top: number } | null>(null)
  const [open, setOpen] = useState<number | null>(null)
  useLayoutEffect(() => {
    if (!parent || !ref.current) return
    const row = parent.getBoundingClientRect()
    const rect = ref.current.getBoundingClientRect()
    const right = row.right + 2
    const left = right + rect.width > window.innerWidth - 4 ? Math.max(4, row.left - rect.width - 2) : right
    const top = Math.max(4, Math.min(row.top - 4, window.innerHeight - rect.height - 4))
    setPlace({ left, top })
  }, [parent])
  useEffect(() => {
    if (place) ref.current?.querySelector<HTMLButtonElement>(".fm-row:not(:disabled)")?.focus()
  }, [place])
  return createPortal(
    <div
      ref={ref}
      className="filter-menu fm-submenu"
      role="menu"
      style={{ left: place?.left ?? 0, top: place?.top ?? 0, visibility: place ? "visible" : "hidden" }}
      onKeyDown={(event) => {
        const rows = [...(ref.current?.querySelectorAll<HTMLButtonElement>(":scope > .fm-row:not(:disabled)") ?? [])]
        const at = rows.indexOf(document.activeElement as HTMLButtonElement)
        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
          event.preventDefault()
          event.stopPropagation()
          rows[(at + (event.key === "ArrowDown" ? 1 : -1) + rows.length) % rows.length]?.focus()
        } else if (event.key === "ArrowLeft") {
          event.preventDefault()
          event.stopPropagation()
          onBack()
        }
      }}
    >
      {items.map((item, i) =>
        item.kind === "separator" ? (
          <div key={i} className="menu-separator" role="separator" />
        ) : !item.run && !item.sub ? (
          <div key={i} className="fm-heading">
            {item.label}
          </div>
        ) : (
          <button
            key={i}
            type="button"
            role={item.checked !== undefined ? "menuitemcheckbox" : "menuitem"}
            aria-checked={item.checked}
            className={cn("fm-row", open === i && "is-open")}
            data-sub={item.sub ? `n${i}` : undefined}
            onPointerEnter={() => setOpen(item.sub ? i : null)}
            onKeyDown={(event) => {
              if (item.sub && event.key === "ArrowRight") {
                event.preventDefault()
                event.stopPropagation()
                setOpen(i)
              }
            }}
            onClick={() => (item.sub ? setOpen(i) : item.run?.())}
          >
            <span className="fm-row-icon">{item.checked ? <Check size={14} strokeWidth={1.8} /> : null}</span>
            {item.icon ? <CfIcon icon={item.icon} /> : null}
            {item.swatch !== undefined ? (
              <span
                className={cn("fm-swatch", item.swatch === "transparent" && "is-none")}
                style={{ background: item.swatch ?? "var(--app-ribbon-text)" }}
                aria-label={item.label || item.swatch || "Automatic"}
              />
            ) : null}
            <span>{item.label}</span>
            {item.sub ? <ChevronRight className="fm-row-more" size={14} strokeWidth={1.6} /> : null}
          </button>
        ),
      )}
      {open !== null && items[open]?.kind === "item" && (items[open] as { sub?: SubItem[] }).sub ? (
        <Submenu
          parent={ref.current?.querySelector<HTMLElement>(`[data-sub="n${open}"]`) ?? null}
          items={(items[open] as { sub: SubItem[] }).sub}
          onBack={() => {
            ref.current?.querySelector<HTMLElement>(`[data-sub="n${open}"]`)?.focus()
            setOpen(null)
          }}
        />
      ) : null}
    </div>,
    document.body,
  )
}

/** A check box that can show "some ticked" (Select All, a year or a month of the date tree). */
function TriBox({
  checked,
  mixed,
  label,
  onChange,
}: {
  checked: boolean
  mixed?: boolean
  label: ReactNode
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

/** (Select All), the values, the dates as a year › month › day tree, and (Blanks). */
function ValueTree({ choices, onChange }: { choices: FilterChoices; onChange: (next: FilterChoices) => void }) {
  const [openYears, setOpenYears] = useState<Set<number>>(new Set())
  const [openMonths, setOpenMonths] = useState<Set<string>>(new Set())
  const all = [...choices.values.map((v) => v.checked), ...choices.dates.map((d) => d.checked)]
  if (choices.blanks) all.push(choices.blanks.checked)
  const every = all.every(Boolean)
  const some = all.some(Boolean)
  const setAll = (checked: boolean) =>
    onChange({
      ...choices,
      values: choices.values.map((v) => ({ ...v, checked })),
      dates: choices.dates.map((d) => ({ ...d, checked })),
      blanks: choices.blanks ? { checked } : null,
    })
  const setDates = (match: (d: FilterChoices["dates"][number]) => boolean, checked: boolean) =>
    onChange({ ...choices, dates: choices.dates.map((d) => (match(d) ? { ...d, checked } : d)) })
  const years = [...new Set(choices.dates.map((d) => d.year))]
  const toggle = <T,>(set: Set<T>, key: T) => {
    const next = new Set(set)
    if (next.has(key)) next.delete(key)
    else next.add(key)
    return next
  }
  return (
    <>
      <TriBox checked={every} mixed={some && !every} label="(Select All)" onChange={setAll} />
      {years.map((year) => {
        const days = choices.dates.filter((d) => d.year === year)
        const months = [...new Set(days.map((d) => d.month))]
        const yearOpen = openYears.has(year)
        return (
          <div key={year} role="treeitem" aria-expanded={yearOpen}>
            <div className="fm-node">
              <button
                type="button"
                className="fm-expand"
                aria-label={yearOpen ? `Collapse ${year}` : `Expand ${year}`}
                onClick={() => setOpenYears(toggle(openYears, year))}
              >
                {yearOpen ? <Minus size={10} strokeWidth={2} /> : <Plus size={10} strokeWidth={2} />}
              </button>
              <TriBox
                checked={days.every((d) => d.checked)}
                mixed={days.some((d) => d.checked) && !days.every((d) => d.checked)}
                label={year}
                onChange={(v) => setDates((d) => d.year === year, v)}
              />
            </div>
            {yearOpen
              ? months.map((month) => {
                  const inMonth = days.filter((d) => d.month === month)
                  const key = `${year}-${month}`
                  const monthOpen = openMonths.has(key)
                  return (
                    <div key={key} className="fm-indent" role="treeitem" aria-expanded={monthOpen}>
                      <div className="fm-node">
                        <button
                          type="button"
                          className="fm-expand"
                          aria-label={monthOpen ? `Collapse ${MONTHS[month - 1]}` : `Expand ${MONTHS[month - 1]}`}
                          onClick={() => setOpenMonths(toggle(openMonths, key))}
                        >
                          {monthOpen ? <Minus size={10} strokeWidth={2} /> : <Plus size={10} strokeWidth={2} />}
                        </button>
                        <TriBox
                          checked={inMonth.every((d) => d.checked)}
                          mixed={inMonth.some((d) => d.checked) && !inMonth.every((d) => d.checked)}
                          label={MONTHS[month - 1]}
                          onChange={(v) => setDates((d) => d.year === year && d.month === month, v)}
                        />
                      </div>
                      {monthOpen
                        ? inMonth.map((d) => (
                            <div key={d.day} className="fm-indent fm-leaf">
                              <TriBox
                                checked={d.checked}
                                label={String(d.day).padStart(2, "0")}
                                onChange={(v) => setDates((x) => x === d, v)}
                              />
                            </div>
                          ))
                        : null}
                    </div>
                  )
                })
              : null}
          </div>
        )
      })}
      {choices.values.map((v, i) => (
        <TriBox
          key={v.text + i}
          checked={v.checked}
          label={v.text}
          onChange={(checked) =>
            onChange({ ...choices, values: choices.values.map((x, k) => (k === i ? { ...x, checked } : x)) })
          }
        />
      ))}
      {choices.blanks ? (
        <TriBox
          checked={choices.blanks.checked}
          label="(Blanks)"
          onChange={(checked) => onChange({ ...choices, blanks: { checked } })}
        />
      ) : null}
    </>
  )
}

/** While searching: the matches, ticked to start with, and "Add current selection to filter". */
function SearchList({
  hits,
  ticked,
  onTick,
  addToFilter,
  onAddToFilter,
}: {
  hits: string[]
  ticked: Set<string>
  onTick: (next: Set<string>) => void
  addToFilter: boolean
  onAddToFilter: (v: boolean) => void
}) {
  if (!hits.length) return <p className="fm-empty">No matches</p>
  const every = hits.every((h) => ticked.has(h))
  return (
    <>
      <TriBox
        checked={every}
        mixed={!every && ticked.size > 0}
        label="(Select All Search Results)"
        onChange={(v) => onTick(new Set(v ? hits : []))}
      />
      <TriBox checked={addToFilter} label="Add current selection to filter" onChange={onAddToFilter} />
      {hits.map((text) => (
        <TriBox
          key={text}
          checked={ticked.has(text)}
          label={text}
          onChange={(v) => {
            const next = new Set(ticked)
            if (v) next.add(text)
            else next.delete(text)
            onTick(next)
          }}
        />
      ))}
    </>
  )
}
