import { columnName, isDateFormat, shownColors, shownIcon, type IconRef, type Range, type SortKey } from "@cells/core"
import { CfIcon } from "@/grid/CfIcon"
import { iconName } from "@/grid/cf-icons"
import { ArrowDown, ArrowUp, Copy, Plus, X } from "lucide-react"
import { useMemo, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** Excel's built-in custom lists (File › Options › Edit Custom Lists). */
export const CUSTOM_LISTS = [
  ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
  ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  [
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
  ],
]

const MAX_LEVELS = 64

type Level = {
  index: number
  on: "values" | "cellColor" | "fontColor" | "icon"
  /** "asc" / "desc", or "list:N" for a custom list, ascending. */
  order: string
  color: string | null
  icon?: IconRef | null
  /** Colour levels: On Top (false) or On Bottom (true). */
  bottom: boolean
}

/** Data › Sort: levels of keys, each a column, what to sort on and in which order; Options for case and direction. */
export function SortDialog({
  controller: c,
  range,
  onClose,
}: {
  controller: SheetController
  range: Range
  onClose: () => void
}) {
  const [header, setHeader] = useState(() => c.data.target(range).header)
  const [byColumns, setByColumns] = useState(false)
  const [matchCase, setMatchCase] = useState(false)
  const [options, setOptions] = useState(false)
  const lines = useMemo(() => {
    const out: number[] = []
    const [from, to] = byColumns ? [range.r0, range.r1] : [range.c0, range.c1]
    for (let i = from; i <= to; i++) out.push(i)
    return out
  }, [range, byColumns])
  const first = (byColumns ? range.c0 : range.r0) + (header ? 1 : 0)
  const last = byColumns ? range.c1 : range.r1
  const cell = (line: number, at: number) => (byColumns ? c.sheet.cells.get(line, at) : c.sheet.cells.get(at, line))

  const nameOf = (line: number) => {
    if (header) {
      const head = cell(line, byColumns ? range.c0 : range.r0)
      if (head?.value !== undefined && head.value !== "") return String(head.value)
    }
    return byColumns ? `Row ${line + 1}` : `Column ${columnName(line)}`
  }
  /** What the line holds, for the Order list's wording. */
  const kindOf = (line: number): "text" | "number" | "date" => {
    for (let at = first; at <= Math.min(last, first + 200); at++) {
      const d = cell(line, at)
      if (d?.value === undefined || d.value === "") continue
      if (typeof d.value !== "number") return "text"
      const code = c.wb.styles.get(d.style ?? 0).numberFormat
      return code && isDateFormat(code) ? "date" : "number"
    }
    return "text"
  }
  const lookAt = (line: number, at: number) => (byColumns ? c.data.look(line, at) : c.data.look(at, line))
  const colorsOf = (line: number, on: Level["on"]) => {
    const found = new Set<string | null>()
    for (let at = first; at <= Math.min(last, first + 5000); at++) {
      const shown = shownColors(c.wb.styles.get(cell(line, at)?.style ?? 0), lookAt(line, at))
      found.add(on === "cellColor" ? shown.fill : shown.font)
    }
    return [...found]
  }
  /** The icons the line shows ("set:index", "" for No Cell Icon), best first. */
  const iconsOf = (line: number) => {
    const found = new Map<string, IconRef | null>()
    for (let at = first; at <= Math.min(last, first + 5000); at++) {
      const icon = shownIcon(lookAt(line, at))
      found.set(icon ? `${icon.set}:${icon.index}` : "", icon)
    }
    return [...found.values()].sort((a, b) => (a ? (b ? b.index - a.index : -1) : b ? 1 : 0))
  }
  const iconKey = (icon: IconRef | null | undefined) => (icon ? `${icon.set}:${icon.index}` : "")

  const active = byColumns ? c.selection.active.row : c.selection.active.col
  const start = lines.includes(active) ? active : lines[0]
  const [levels, setLevels] = useState<Level[]>([
    { index: start, on: "values", order: "asc", color: null, bottom: false },
  ])
  const [current, setCurrent] = useState(0)
  const update = (i: number, patch: Partial<Level>) =>
    setLevels((list) => list.map((l, k) => (k === i ? { ...l, ...patch } : l)))
  const add = () => {
    if (levels.length >= MAX_LEVELS) return
    const used = new Set(levels.map((l) => l.index))
    const next = lines.find((l) => !used.has(l)) ?? lines[0]
    setLevels([...levels, { index: next, on: "values", order: "asc", color: null, bottom: false }])
    setCurrent(levels.length)
  }
  const move = (dir: -1 | 1) => {
    const to = current + dir
    if (to < 0 || to >= levels.length) return
    const list = [...levels]
    ;[list[current], list[to]] = [list[to], list[current]]
    setLevels(list)
    setCurrent(to)
  }

  const accept = () => {
    const keys: SortKey[] = levels.map((l) => {
      if (l.on === "icon") return { index: l.index, on: "icon", icon: l.icon ?? null, descending: l.bottom }
      if (l.on !== "values") return { index: l.index, on: l.on, color: l.color, descending: l.bottom }
      if (l.order.startsWith("list:")) return { index: l.index, list: CUSTOM_LISTS[Number(l.order.slice(5))] }
      return { index: l.index, descending: l.order === "desc" }
    })
    if (c.data.sort(range, { keys, header, matchCase, byColumns })) onClose()
  }

  const orderNames = (line: number) => {
    const kind = kindOf(line)
    return kind === "number"
      ? ["Smallest to Largest", "Largest to Smallest"]
      : kind === "date"
        ? ["Oldest to Newest", "Newest to Oldest"]
        : ["A to Z", "Z to A"]
  }

  return (
    <Dialog title="Sort" width={640} onCancel={onClose} onAccept={accept}>
      <div className="sort-dialog">
        <div className="sort-toolbar">
          <button type="button" className="dialog-btn sort-tool" onClick={add} disabled={levels.length >= MAX_LEVELS}>
            <Plus size={14} /> Add Level
          </button>
          <button
            type="button"
            className="dialog-btn sort-tool"
            disabled={levels.length < 2}
            onClick={() => {
              setLevels(levels.filter((_, k) => k !== current))
              setCurrent(Math.max(0, current - 1))
            }}
          >
            <X size={14} /> Delete Level
          </button>
          <button
            type="button"
            className="dialog-btn sort-tool"
            disabled={levels.length >= MAX_LEVELS}
            onClick={() => {
              const list = [...levels]
              list.splice(current + 1, 0, { ...levels[current] })
              setLevels(list)
              setCurrent(current + 1)
            }}
          >
            <Copy size={14} /> Copy Level
          </button>
          <button
            type="button"
            className="dialog-btn sort-tool"
            aria-label="Move Up"
            title="Move Up"
            disabled={current === 0}
            onClick={() => move(-1)}
          >
            <ArrowUp size={14} />
          </button>
          <button
            type="button"
            className="dialog-btn sort-tool"
            aria-label="Move Down"
            title="Move Down"
            disabled={current === levels.length - 1}
            onClick={() => move(1)}
          >
            <ArrowDown size={14} />
          </button>
          <button type="button" className="dialog-btn" onClick={() => setOptions(true)}>
            Options...
          </button>
          <label className="dialog-radio sort-headers">
            <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
            My data has headers
          </label>
        </div>
        <div className="sort-levels" role="table" aria-label="Sort levels">
          <div className="sort-level sort-level-head" role="row">
            <span role="columnheader" />
            <span role="columnheader">{byColumns ? "Row" : "Column"}</span>
            <span role="columnheader">Sort On</span>
            <span role="columnheader">Order</span>
          </div>
          {levels.map((level, i) => {
            const colors = level.on === "values" || level.on === "icon" ? [] : colorsOf(level.index, level.on)
            const icons = level.on === "icon" ? iconsOf(level.index) : []
            return (
              <div
                key={i}
                role="row"
                className={i === current ? "sort-level is-current" : "sort-level"}
                aria-selected={i === current}
                onClick={() => setCurrent(i)}
                onFocus={() => setCurrent(i)}
              >
                <span className="sort-level-label">{i === 0 ? "Sort by" : "Then by"}</span>
                <select
                  className="dialog-field"
                  aria-label={`${i === 0 ? "Sort by" : "Then by"} ${byColumns ? "row" : "column"}`}
                  value={level.index}
                  onChange={(e) => update(i, { index: Number(e.target.value) })}
                >
                  {lines.map((line) => (
                    <option key={line} value={line}>
                      {nameOf(line)}
                    </option>
                  ))}
                </select>
                <select
                  className="dialog-field"
                  aria-label="Sort On"
                  value={level.on}
                  onChange={(e) => {
                    const on = e.target.value as Level["on"]
                    if (on === "icon") update(i, { on, color: null, icon: iconsOf(level.index)[0] ?? null })
                    else update(i, { on, color: on === "values" ? null : (colorsOf(level.index, on)[0] ?? null) })
                  }}
                >
                  <option value="values">Cell Values</option>
                  <option value="cellColor">Cell Color</option>
                  <option value="fontColor">Font Color</option>
                  <option value="icon">Conditional Formatting Icon</option>
                </select>
                {level.on === "values" ? (
                  <select
                    className="dialog-field"
                    aria-label="Order"
                    value={level.order}
                    onChange={(e) => update(i, { order: e.target.value })}
                  >
                    <option value="asc">{orderNames(level.index)[0]}</option>
                    <option value="desc">{orderNames(level.index)[1]}</option>
                    {CUSTOM_LISTS.map((list, k) => (
                      <option key={k} value={`list:${k}`}>
                        {list.join(", ")}
                      </option>
                    ))}
                  </select>
                ) : (
                  <span className="sort-color">
                    {level.on === "icon" ? (
                      <>
                        <select
                          className="dialog-field"
                          aria-label="Icon"
                          value={iconKey(level.icon)}
                          onChange={(e) =>
                            update(i, { icon: icons.find((x) => iconKey(x) === e.target.value) ?? null })
                          }
                        >
                          {icons.map((icon) => (
                            <option key={iconKey(icon)} value={iconKey(icon)}>
                              {iconName(icon)}
                            </option>
                          ))}
                        </select>
                        {level.icon ? (
                          <CfIcon icon={level.icon} />
                        ) : (
                          <span className="sort-swatch" aria-hidden data-empty="" />
                        )}
                      </>
                    ) : (
                      <>
                        <select
                          className="dialog-field"
                          aria-label="Color"
                          value={level.color ?? ""}
                          onChange={(e) => update(i, { color: e.target.value || null })}
                        >
                          {colors.map((color) => (
                            <option key={color ?? "none"} value={color ?? ""}>
                              {color ?? (level.on === "cellColor" ? "No Cell Color" : "Automatic")}
                            </option>
                          ))}
                        </select>
                        <span
                          className="sort-swatch"
                          aria-hidden
                          style={{ background: level.color ?? "transparent" }}
                          data-empty={level.color ? undefined : ""}
                        />
                      </>
                    )}
                    <select
                      className="dialog-field"
                      aria-label="Position"
                      value={level.bottom ? "bottom" : "top"}
                      onChange={(e) => update(i, { bottom: e.target.value === "bottom" })}
                    >
                      <option value="top">On Top</option>
                      <option value="bottom">On Bottom</option>
                    </select>
                  </span>
                )}
              </div>
            )
          })}
        </div>
      </div>
      {options ? (
        <SortOptionsDialog
          matchCase={matchCase}
          byColumns={byColumns}
          onClose={() => setOptions(false)}
          onAccept={(o) => {
            setMatchCase(o.matchCase)
            if (o.byColumns !== byColumns) {
              setByColumns(o.byColumns)
              const lines2 = o.byColumns ? range.r0 : range.c0
              setLevels([{ index: lines2, on: "values", order: "asc", color: null, bottom: false }])
              setCurrent(0)
              setHeader(false)
            }
            setOptions(false)
          }}
        />
      ) : null}
    </Dialog>
  )
}

function SortOptionsDialog({
  matchCase: initialCase,
  byColumns: initialColumns,
  onAccept,
  onClose,
}: {
  matchCase: boolean
  byColumns: boolean
  onAccept: (o: { matchCase: boolean; byColumns: boolean }) => void
  onClose: () => void
}) {
  const [matchCase, setMatchCase] = useState(initialCase)
  const [byColumns, setByColumns] = useState(initialColumns)
  return (
    <Dialog title="Sort Options" width={260} onCancel={onClose} onAccept={() => onAccept({ matchCase, byColumns })}>
      <div className="sort-options">
        <label className="dialog-radio">
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          Case sensitive
        </label>
        <fieldset className="tiw-fieldset">
          <legend>Orientation</legend>
          <label className="dialog-radio">
            <input type="radio" checked={!byColumns} onChange={() => setByColumns(false)} />
            Sort top to bottom
          </label>
          <label className="dialog-radio">
            <input type="radio" checked={byColumns} onChange={() => setByColumns(true)} />
            Sort left to right
          </label>
        </fieldset>
      </div>
    </Dialog>
  )
}

/**
 * Excel's Sort Warning (and Remove Duplicates Warning): data next to the selection won't be included unless the
 * selection grows.
 */
export function SortWarningDialog({
  onChoose,
  onClose,
  kind = "sort",
}: {
  onChoose: (expand: boolean) => void
  onClose: () => void
  kind?: "sort" | "duplicates"
}) {
  const [expand, setExpand] = useState(true)
  const sort = kind === "sort"
  return (
    <Dialog
      title={sort ? "Sort Warning" : "Remove Duplicates Warning"}
      width={420}
      acceptLabel={sort ? "Sort" : "Remove Duplicates..."}
      onCancel={onClose}
      onAccept={() => onChoose(expand)}
    >
      <p className="dialog-message">
        {sort
          ? "Cells found data next to your selection. Since you have not selected this data, it will not be sorted."
          : "Cells found data next to your selection. Since you have not selected this data, it will not be considered."}
      </p>
      <p className="dialog-message sort-warning-ask">What do you want to do?</p>
      <div className="dialog-group" role="radiogroup" aria-label="What do you want to do?">
        <label className="dialog-radio">
          <input type="radio" checked={expand} onChange={() => setExpand(true)} />
          Expand the selection
        </label>
        <label className="dialog-radio">
          <input type="radio" checked={!expand} onChange={() => setExpand(false)} />
          Continue with the current selection
        </label>
      </div>
    </Dialog>
  )
}
