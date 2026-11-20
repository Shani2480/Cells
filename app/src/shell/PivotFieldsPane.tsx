import {
  pivotFieldAreas,
  pivotFieldName,
  VALUES_FIELD,
  withFieldIn,
  withFieldRemoved,
  withFieldTicked,
  withoutField,
  withValueMoved,
  type PivotArea,
  type PivotCache,
  type PivotTable,
} from "@cells/core"
import { ChevronDown, Columns3, Filter, Rows3, Sigma, X } from "lucide-react"
import { useState, type CSSProperties, type DragEvent, type ReactNode } from "react"
import type { SheetController } from "@/sheet/controller"

/** A field chip in one of the four areas: its field and its place there. */
type Chip = { area: PivotArea; f: number; index: number }

/** What a drag carries: a field from the list, or a chip from an area. */
type Drag = { from: "list"; f: number } | ({ from: "area" } & Chip)

const AREAS: { area: PivotArea; label: string; icon: ReactNode }[] = [
  { area: "filters", label: "Filters", icon: <Filter size={14} strokeWidth={1.75} /> },
  { area: "cols", label: "Columns", icon: <Columns3 size={14} strokeWidth={1.75} /> },
  { area: "rows", label: "Rows", icon: <Rows3 size={14} strokeWidth={1.75} /> },
  { area: "values", label: "Values", icon: <Sigma size={14} strokeWidth={1.75} /> },
]

const ADD_TO: { area: PivotArea; label: string }[] = [
  { area: "filters", label: "Add to Report Filter" },
  { area: "rows", label: "Add to Row Labels" },
  { area: "cols", label: "Add to Column Labels" },
  { area: "values", label: "Add to Values" },
]

const MOVE_TO: { area: PivotArea; label: string }[] = [
  { area: "filters", label: "Move to Report Filter" },
  { area: "rows", label: "Move to Row Labels" },
  { area: "cols", label: "Move to Column Labels" },
  { area: "values", label: "Move to Values" },
]

/** The fields in an area, in order (the Σ Values field where it sits among rows or columns). */
function chips(table: PivotTable, area: PivotArea): number[] {
  if (area === "rows") return table.rows
  if (area === "cols") return table.cols
  if (area === "filters") return table.pages.map((p) => p.field)
  return table.values.map((v) => v.field)
}

/** The table with chip `c` moved to `area` at `index` (the end when undefined). */
export function movedChip(table: PivotTable, cache: PivotCache, c: Chip, area: PivotArea, index?: number): PivotTable {
  if (c.area === area) {
    const to = index === undefined ? chips(table, area).length - 1 : index > c.index ? index - 1 : index
    if (to === c.index) return table
    if (area === "values") return withValueMoved(table, c.index, to)
    return withFieldIn(table, cache, c.f, area, to)
  }
  if (c.f === VALUES_FIELD && area !== "rows" && area !== "cols") return table
  // Out of Values first (that value field only); Filters, Columns and Rows give the field up on their own.
  const base = c.area === "values" ? withoutField(table, c.f, "values", c.index) : table
  return withFieldIn(base, cache, c.f, area, index)
}

/**
 * The PivotTable Fields pane: the source's fields, ticked when the PivotTable uses them, and the Filters, Columns,
 * Rows and Values areas below. Ticking adds a field where Excel would (numbers to Values, the rest to Rows), dragging
 * puts it in an area or moves it within one, and each chip's menu moves or removes it. Every change is one undo step.
 */
export function PivotFieldsPane({
  controller,
  onDone,
  onDialog,
}: {
  controller: SheetController
  onDone: () => void
  /** Opens Value Field Settings ("pivotValue:i") or Field Settings ("pivotField:f"). */
  onDialog?: (name: string) => void
}) {
  const pivots = controller.pivots
  const [search, setSearch] = useState("")
  const [menu, setMenu] = useState<string | null>(null)
  // Menus sit on the window (the list and the areas scroll and would cut them off), below their row or above it.
  const [place, setPlace] = useState<CSSProperties>({})
  const openMenu = (key: string, row: Element | null) => {
    if (menu === key) return setMenu(null)
    const r = row?.getBoundingClientRect()
    if (r) {
      const right = Math.max(4, window.innerWidth - r.right)
      setPlace(
        window.innerHeight - r.bottom > 280
          ? { position: "fixed", top: r.bottom, bottom: "auto", left: "auto", right }
          : { position: "fixed", top: "auto", bottom: window.innerHeight - r.top, left: "auto", right },
      )
    }
    setMenu(key)
  }
  const [over, setOver] = useState<string | null>(null)
  // Fingers can't drag here (HTML drag and drop is mouse-only in most browsers): the lead says to use the menus.
  const [coarse] = useState(() => typeof matchMedia === "function" && matchMedia("(pointer: coarse)").matches)
  const at = pivots.paneTable()
  if (!at) return null
  const { table, cache, waiting } = at
  const close = () => {
    pivots.setFieldsOpen(false)
    onDone()
  }
  const change = (make: (t: PivotTable, c: PivotCache) => PivotTable) => {
    setMenu(null)
    pivots.paneChange(make)
  }
  const name = (f: number) => (f === VALUES_FIELD ? "Σ Values" : pivotFieldName(table, cache, f))
  const query = search.trim().toLowerCase()
  const fields = cache.fields
    .map((_, f) => f)
    .filter((f) => !query || pivotFieldName(table, cache, f).toLowerCase().includes(query))

  let drag: Drag | null = null
  const startDrag = (e: DragEvent, d: Drag) => {
    drag = d
    e.dataTransfer.setData("application/x-cells-pivot-field", JSON.stringify(d))
    e.dataTransfer.effectAllowed = "move"
  }
  const readDrag = (e: DragEvent): Drag | null => {
    try {
      return drag ?? (JSON.parse(e.dataTransfer.getData("application/x-cells-pivot-field")) as Drag)
    } catch {
      return null
    }
  }
  const drop = (e: DragEvent, area: PivotArea, index?: number) => {
    e.preventDefault()
    e.stopPropagation()
    setOver(null)
    const d = readDrag(e)
    if (!d) return
    if (d.from === "list") change((t, c) => withFieldIn(t, c, d.f, area, index))
    else change((t, c) => movedChip(t, c, d, area, index))
  }
  const dragOver = (e: DragEvent, key: string) => {
    e.preventDefault()
    e.dataTransfer.dropEffect = "move"
    if (over !== key) setOver(key)
  }

  const chipMenu = (c: Chip, count: number) => {
    const values = c.area === "values"
    const move = (index: number) => change((t, k) => movedChip(t, k, c, c.area, index))
    return (
      <div className="pivot-chip-menu" role="menu" aria-label={name(c.f)} style={place}>
        <button type="button" role="menuitem" disabled={c.index === 0} onClick={() => move(c.index - 1)}>
          Move Up
        </button>
        <button type="button" role="menuitem" disabled={c.index === count - 1} onClick={() => move(c.index + 2)}>
          Move Down
        </button>
        <button type="button" role="menuitem" disabled={c.index === 0} onClick={() => move(0)}>
          Move to Beginning
        </button>
        <button type="button" role="menuitem" disabled={c.index === count - 1} onClick={() => move(count)}>
          Move to End
        </button>
        <hr />
        {MOVE_TO.map((m) => (
          <button
            key={m.area}
            type="button"
            role="menuitem"
            disabled={m.area === c.area || (c.f === VALUES_FIELD && (m.area === "filters" || m.area === "values"))}
            onClick={() => change((t, k) => movedChip(t, k, c, m.area))}
          >
            {m.label}
          </button>
        ))}
        <hr />
        <button
          type="button"
          role="menuitem"
          disabled={c.f === VALUES_FIELD}
          onClick={() => change((t) => withoutField(t, c.f, c.area, values ? c.index : undefined))}
        >
          Remove Field
        </button>
        <hr />
        <button
          type="button"
          role="menuitem"
          disabled={c.f === VALUES_FIELD}
          onClick={() => {
            setMenu(null)
            onDialog?.(values ? `pivotValue:${c.index}` : `pivotField:${c.f}`)
          }}
        >
          {values ? "Value Field Settings..." : "Field Settings..."}
        </button>
      </div>
    )
  }

  return (
    <aside
      className="comments-pane pivot-fields-pane"
      aria-label="PivotTable Fields"
      onPointerDown={(e) => {
        // A press outside the open menu (and its button) closes it.
        const t = e.target as HTMLElement
        if (menu && !t.closest(".pivot-chip-menu, .pivot-chip-button")) setMenu(null)
      }}
      onKeyDown={(e) => {
        if (e.key === "Escape") {
          e.preventDefault()
          if (menu) setMenu(null)
          else close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>PivotTable Fields</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <p className="pivot-fields-lead">Choose fields to add to report:</p>
      <input
        className="dialog-field pivot-fields-search"
        type="search"
        placeholder="Search"
        aria-label="Search"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      <ul
        className="pivot-fields-list"
        aria-label="Fields"
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          // A chip dropped back on the list leaves its area.
          e.preventDefault()
          const d = readDrag(e)
          if (d?.from === "area" && d.f !== VALUES_FIELD)
            change((t) => withoutField(t, d.f, d.area, d.area === "values" ? d.index : undefined))
        }}
      >
        {fields.map((f) => {
          const used = pivotFieldAreas(table, f).length > 0
          return (
            <li
              key={f}
              className="pivot-field"
              draggable
              onDragStart={(e) => startDrag(e, { from: "list", f })}
              onContextMenu={(e) => {
                e.preventDefault()
                setMenu(null)
                openMenu(`field:${f}`, e.currentTarget)
              }}
            >
              <label className="dialog-radio">
                <input
                  type="checkbox"
                  checked={used}
                  onChange={() => change((t, c) => (used ? withFieldRemoved(t, f) : withFieldTicked(t, c, f)))}
                />
                <span className={used ? "is-used" : undefined}>{pivotFieldName(table, cache, f)}</span>
              </label>
              <button
                type="button"
                className="pivot-chip-button"
                aria-label={`${pivotFieldName(table, cache, f)} field menu`}
                aria-haspopup="menu"
                aria-expanded={menu === `field:${f}`}
                onClick={(e) => openMenu(`field:${f}`, e.currentTarget.closest("li"))}
              >
                <ChevronDown size={14} strokeWidth={1.75} />
              </button>
              {menu === `field:${f}` ? (
                // Excel's menu on a field of the list: where to add it.
                <div className="pivot-chip-menu" role="menu" aria-label={pivotFieldName(table, cache, f)} style={place}>
                  {ADD_TO.map((m) => (
                    <button
                      key={m.area}
                      type="button"
                      role="menuitem"
                      disabled={m.area !== "values" && pivotFieldAreas(table, f).includes(m.area)}
                      onClick={() => change((t, c) => withFieldIn(t, c, f, m.area))}
                    >
                      {m.label}
                    </button>
                  ))}
                </div>
              ) : null}
            </li>
          )
        })}
      </ul>
      <p className="pivot-fields-lead">
        {coarse ? "Move fields between areas below with their menus:" : "Drag fields between areas below:"}
      </p>
      <div className="pivot-areas">
        {AREAS.map(({ area, label, icon }) => {
          const list = chips(table, area)
          return (
            <section
              key={area}
              className={`pivot-area${over === area ? " is-over" : ""}`}
              aria-label={label}
              onDragOver={(e) => dragOver(e, area)}
              onDragLeave={() => setOver(null)}
              onDrop={(e) => drop(e, area)}
            >
              <h3>
                {icon}
                {label}
              </h3>
              <ul>
                {list.map((f, index) => {
                  const key = `${area}:${index}`
                  const chip: Chip = { area, f, index }
                  const text = area === "values" ? (table.values[index]?.name ?? name(f)) : name(f)
                  return (
                    <li
                      key={key}
                      className={`pivot-chip${over === key ? " is-over" : ""}`}
                      draggable
                      onDragStart={(e) => startDrag(e, { from: "area", ...chip })}
                      onDragOver={(e) => dragOver(e, key)}
                      onDrop={(e) => drop(e, area, index)}
                    >
                      <span className="pivot-chip-name">{text}</span>
                      <button
                        type="button"
                        className="pivot-chip-button"
                        aria-label={`${text} options`}
                        aria-haspopup="menu"
                        aria-expanded={menu === key}
                        onClick={(e) => openMenu(key, e.currentTarget.closest("li"))}
                      >
                        <ChevronDown size={14} strokeWidth={1.75} />
                      </button>
                      {menu === key ? chipMenu(chip, list.length) : null}
                    </li>
                  )
                })}
              </ul>
            </section>
          )
        })}
      </div>
      <footer className="pivot-fields-foot">
        <label className="dialog-radio">
          <input type="checkbox" checked={pivots.deferred} onChange={(e) => pivots.setDeferred(e.target.checked)} />
          <span>Defer Layout Update</span>
        </label>
        <button type="button" className="dialog-btn" disabled={!waiting} onClick={() => pivots.update()}>
          Update
        </button>
      </footer>
    </aside>
  )
}
