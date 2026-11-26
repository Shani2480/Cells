import { readFileSync } from "node:fs"
import { resolve } from "node:path"
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { buildRegistry } from "@/ribbon/commands"
import { useSettings } from "@/settings/settings"
import { DEFAULT_SETTINGS } from "@/settings/shared"
import { FIRST_HINT_MS, FeatureHints, NEXT_HINT_MS } from "./FeatureHints"
import { FEATURE_HINTS, nextHint } from "./feature-hints"

/** A control as the browser lays it out (jsdom gives every element an empty box). */
function control(label: string) {
  const button = document.createElement("button")
  button.setAttribute("aria-label", label)
  button.getBoundingClientRect = () =>
    ({ left: 100, top: 40, right: 160, bottom: 100, width: 60, height: 60 }) as DOMRect
  document.body.append(button)
  return button
}

beforeEach(() => {
  vi.useFakeTimers()
  useSettings.setState({ ...DEFAULT_SETTINGS, loaded: true, setupDone: true })
})
afterEach(() => {
  cleanup()
  document.body.innerHTML = ""
  vi.useRealTimers()
  useSettings.setState({ ...DEFAULT_SETTINGS, loaded: false })
})

describe("feature hints", () => {
  it("introduce only special features, each as news, pointing at a control Cells has", () => {
    const ids = FEATURE_HINTS.map((hint) => hint.id)
    expect(new Set(ids).size).toBe(ids.length)
    const labels = new Set(
      buildRegistry()
        .all()
        .map((c) => c.label),
    )
    // Controls outside the ribbon, named in their own components.
    const elsewhere = [
      readFileSync(resolve(__dirname, "../grid/QuickAnalysis.tsx"), "utf8"),
      readFileSync(resolve(__dirname, "../shell/PivotFieldsPane.tsx"), "utf8"),
      readFileSync(resolve(__dirname, "../shell/WorkbookSwitcher.tsx"), "utf8"),
    ].join("\n")
    for (const hint of FEATURE_HINTS) {
      expect(hint.title.startsWith("New: "), hint.id).toBe(true)
      expect(labels.has(hint.target) || elsewhere.includes(`aria-label="${hint.target}"`), hint.target).toBe(true)
    }
    // Everyday commands get none.
    const targets = FEATURE_HINTS.map((hint) => hint.target)
    for (const everyday of ["Bold", "Sort A to Z", "Save", "Insert Chart", "Filter"])
      expect(targets).not.toContain(everyday)
  })

  it("pick the first one not dismissed whose control is on screen", () => {
    const shown = (target: string) => target === "AI Tools" || target === "New Window"
    expect(nextHint(FEATURE_HINTS, [], shown)?.id).toBe("ai-tools")
    expect(nextHint(FEATURE_HINTS, ["ai-tools"], shown)?.id).toBe("new-window")
    expect(nextHint(FEATURE_HINTS, ["ai-tools", "new-window"], shown)).toBeNull()
  })

  it("show one at a time after a while, go for good on a tap, and wait before the next", () => {
    control("Flash Fill")
    control("New Window")
    render(<FeatureHints />)
    expect(screen.queryByRole("dialog")).toBeNull()
    act(() => vi.advanceTimersByTime(FIRST_HINT_MS + 600))
    expect(screen.getAllByRole("dialog")).toHaveLength(1)
    expect(screen.getByText("New: Flash Fill as you type")).toBeTruthy()

    fireEvent.click(screen.getByRole("dialog"))
    expect(screen.queryByRole("dialog")).toBeNull()
    expect(useSettings.getState().hintsDismissed).toEqual(["flash-fill-live"])

    act(() => vi.advanceTimersByTime(NEXT_HINT_MS - 1000))
    expect(screen.queryByRole("dialog")).toBeNull()
    act(() => vi.advanceTimersByTime(1600))
    expect(screen.getByText("New: New Window")).toBeTruthy()
  })

  it("go when the feature is used, and wait for a control that is not on screen", () => {
    const flash = control("Flash Fill")
    render(<FeatureHints />)
    act(() => vi.advanceTimersByTime(FIRST_HINT_MS + 600))
    fireEvent.click(flash)
    expect(screen.queryByRole("dialog")).toBeNull()
    expect(useSettings.getState().hintsDismissed).toEqual(["flash-fill-live"])
    act(() => vi.advanceTimersByTime(NEXT_HINT_MS + 600))
    expect(screen.queryByRole("dialog")).toBeNull()
  })

  it("wait while a dialog or menu is open, and skip disabled controls", () => {
    const flash = control("Flash Fill")
    flash.setAttribute("aria-disabled", "true")
    control("New Window")
    const menu = document.createElement("div")
    menu.setAttribute("role", "menu")
    menu.getClientRects = () => [{}] as unknown as DOMRectList
    document.body.append(menu)
    render(<FeatureHints />)
    act(() => vi.advanceTimersByTime(FIRST_HINT_MS + 600))
    expect(screen.queryByRole("dialog")).toBeNull()
    menu.remove()
    act(() => vi.advanceTimersByTime(600))
    expect(screen.getByText("New: New Window")).toBeTruthy()
  })

  it("can be turned off, and shown again from the start", () => {
    control("Flash Fill")
    useSettings.setState({ hintsEnabled: false, hintsDismissed: ["flash-fill-live"] })
    render(<FeatureHints />)
    act(() => vi.advanceTimersByTime(FIRST_HINT_MS + 600))
    expect(screen.queryByRole("dialog")).toBeNull()
    act(() => void useSettings.setState({ hintsEnabled: true, hintsDismissed: [] }))
    act(() => vi.advanceTimersByTime(600))
    expect(screen.getByText("New: Flash Fill as you type")).toBeTruthy()
  })

  it("wait until the first-run setup is done", () => {
    control("Flash Fill")
    useSettings.setState({ setupDone: false })
    render(<FeatureHints />)
    act(() => vi.advanceTimersByTime(FIRST_HINT_MS + 600))
    expect(screen.queryByText("New: Flash Fill as you type")).toBeNull()
  })
})
