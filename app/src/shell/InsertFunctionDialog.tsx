import { FUNCTION_CATEGORIES, FUNCTION_INFO, functionsIn, type FunctionCategory } from "@cells/core"
import { useEffect, useMemo, useRef, useState } from "react"
import { Dialog } from "./Dialog"

const RECENT_KEY = "cells.recent-functions"
/** Excel's starting "Most Recently Used" list (the ones Cells has). */
const DEFAULT_RECENT = ["SUM", "AVERAGE", "IF", "COUNT", "MAX", "SIN", "SUMIF", "PMT", "STDEV"]

export function recentFunctions(): string[] {
  try {
    const saved = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "null") as unknown
    if (Array.isArray(saved)) return saved.filter((n): n is string => typeof n === "string" && FUNCTION_INFO.has(n))
  } catch {
    // Storage blocked or garbled: start from Excel's list.
  }
  return DEFAULT_RECENT
}

export function rememberFunction(name: string) {
  const next = [name, ...recentFunctions().filter((n) => n !== name)].slice(0, 10)
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(next))
  } catch {
    // Not kept; the list stays as it was.
  }
}

/** Search: every word must appear in the name or the description (Excel's Go ranks by its own index). */
export function searchFunctions(query: string): string[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (!words.length) return []
  return [...FUNCTION_INFO.values()]
    .filter((f) => words.every((w) => f.name.toLowerCase().includes(w) || f.description.toLowerCase().includes(w)))
    .map((f) => f.name)
}

type Category = "Most Recently Used" | "All" | FunctionCategory | "Recommended"

/**
 * Insert Function (fx, Shift+F3): search or pick a category, pick a function, OK writes it into the cell. Excel's
 * Function Arguments dialog that follows is not built; the argument tooltip shows while typing instead.
 */
export function InsertFunctionDialog({ onInsert, onClose }: { onInsert: (name: string) => void; onClose: () => void }) {
  const [query, setQuery] = useState("")
  const [category, setCategory] = useState<Category>("Most Recently Used")
  const [found, setFound] = useState<string[] | null>(null)
  const list = useMemo(() => {
    if (category === "Recommended") return found ?? []
    if (category === "Most Recently Used") return recentFunctions()
    return functionsIn(category === "All" ? null : category)
  }, [category, found])
  const [picked, setPicked] = useState<string | null>(null)
  const current = picked && list.includes(picked) ? picked : (list[0] ?? null)
  const info = current ? FUNCTION_INFO.get(current) : undefined
  const listRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    listRef.current?.querySelector(".is-selected")?.scrollIntoView?.({ block: "nearest" })
  }, [current])
  const accept = () => {
    if (!current) return
    rememberFunction(current)
    onInsert(current)
  }
  const go = () => {
    setFound(searchFunctions(query))
    setCategory("Recommended")
    setPicked(null)
  }
  return (
    <Dialog title="Insert Function" onCancel={onClose} onAccept={accept} width={460}>
      <label className="dialog-label" htmlFor="fn-search">
        <u>S</u>earch for a function:
      </label>
      <div className="fn-search-row">
        <input
          id="fn-search"
          className="dialog-field"
          data-autofocus
          placeholder="Type a brief description of what you want to do and then click Go"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && query.trim()) {
              e.preventDefault()
              e.stopPropagation()
              go()
            }
          }}
        />
        <button type="button" className="dialog-btn" disabled={!query.trim()} onClick={go}>
          <u>G</u>o
        </button>
      </div>
      <label className="dialog-label" htmlFor="fn-category">
        Or select a <u>c</u>ategory:
      </label>
      <select
        id="fn-category"
        className="dialog-field"
        value={category}
        onChange={(e) => {
          setCategory(e.target.value as Category)
          setPicked(null)
        }}
      >
        {found ? <option>Recommended</option> : null}
        <option>Most Recently Used</option>
        <option>All</option>
        {FUNCTION_CATEGORIES.map((c) => (
          <option key={c}>{c}</option>
        ))}
      </select>
      <span className="dialog-label">
        Select a functio<u>n</u>:
      </span>
      <div ref={listRef} className="dialog-list fn-list" role="listbox" aria-label="Select a function">
        {list.map((name) => (
          <button
            type="button"
            role="option"
            key={name}
            aria-selected={name === current}
            className={`dialog-list-item${name === current ? " is-selected" : ""}`}
            onClick={() => setPicked(name)}
            onDoubleClick={() => {
              rememberFunction(name)
              onInsert(name)
            }}
          >
            {name}
          </button>
        ))}
        {!list.length ? <span className="fn-empty">No functions match your search.</span> : null}
      </div>
      {info ? (
        <div className="fn-info" aria-live="polite">
          <strong>
            {info.name}({info.args.join(",")})
          </strong>
          <span>{info.description}</span>
        </div>
      ) : null}
    </Dialog>
  )
}
