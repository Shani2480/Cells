import type { Platform } from "@cells/platform"
import { act, cleanup, fireEvent, render, screen, within } from "@testing-library/react"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { FeatureHints } from "@/hints/FeatureHints"
import { FEATURE_HINTS } from "@/hints/feature-hints"
import { pointAt, usePointer } from "@/hints/pointer"
import { useSettings } from "@/settings/settings"
import { DEFAULT_SETTINGS } from "@/settings/shared"
import { PlatformContext } from "@/platform/context"
import { buildRegistry } from "@/ribbon/commands"
import type { RibbonEnv } from "@/ribbon/registry"
import { SheetController } from "@/sheet/controller"
import { FeedbackDialog } from "./FeedbackDialog"
import { issueBody, issueUrl, systemLine } from "./feedback"
import { FEATURE_TOPICS, keysText, placeOf, SHORTCUTS, searchHelp, tabOf } from "./help-content"
import { HelpPane } from "./HelpPane"

/** A ribbon button as the browser lays it out (jsdom gives every element an empty box). */
function control(label: string) {
  const button = document.createElement("button")
  button.setAttribute("aria-label", label)
  button.getBoundingClientRect = () =>
    ({ left: 100, top: 40, right: 160, bottom: 100, width: 60, height: 60 }) as DOMRect
  document.body.append(button)
  return button
}

afterEach(() => {
  cleanup()
  document.body.innerHTML = ""
  usePointer.setState({ pointed: null })
  vi.useRealTimers()
})

describe("what Help knows", () => {
  it("has every feature the hints introduce, each with where it is", () => {
    expect(FEATURE_TOPICS.map((t) => t.id)).toEqual(FEATURE_HINTS.map((h) => h.id))
    for (const topic of FEATURE_TOPICS) {
      expect(topic.title).not.toMatch(/^New:/)
      // A ribbon command on a tab, or words saying where it is.
      if (topic.command) expect(tabOf(topic.command), topic.id).toBeDefined()
      else expect(topic.where, topic.id).toBeTruthy()
    }
    expect(placeOf("QuickCode")).toBe("Data › Data Tools")
    expect(placeOf("WindowNew")).toBe("View › Window")
  })

  it("finds features and shortcuts by any of their words", () => {
    const names = searchHelp("split names")
    expect(names.topics.map((t) => t.id)).toEqual(["flash-fill-live"])
    const undo = searchHelp("undo")
    expect(undo.topics).toEqual([])
    expect(undo.shortcuts.flatMap((g) => g.shortcuts.map((s) => s.keys))).toEqual(["Ctrl+Z / Ctrl+Y"])
    expect(searchHelp("  ").shortcuts).toBe(SHORTCUTS)
  })

  it("writes the keys the Mac way on a Mac", () => {
    expect(keysText("Ctrl+Shift+L", true)).toBe("⌘⇧L")
    expect(keysText("Ctrl+Shift+L", false)).toBe("Ctrl+Shift+L")
  })
})

describe("the Help tab's commands", () => {
  it("Help and Show Training open the pane; Contact Support and Feedback their dialog", () => {
    const c = new SheetController()
    const dialog = vi.fn()
    const env = { c, dialog } as unknown as RibbonEnv
    const registry = buildRegistry()
    for (const id of ["Help", "ContactUs", "OfficeFeedbackHelpTab", "ShowTraining", "WhatsNewRecentUpdates"])
      expect(registry.isBuilt(id), id).toBe(true)
    registry.get("Help")!.run!(env)
    expect(c.helpPane).toBe("help")
    expect(registry.get("Help")!.pressed!(env)).toBe(true)
    registry.get("ShowTraining")!.run!(env)
    expect(c.helpPane).toBe("training")
    registry.get("Help")!.run!(env)
    registry.get("Help")!.run!(env)
    expect(c.helpPane).toBeNull()
    registry.get("ContactUs")!.run!(env)
    registry.get("OfficeFeedbackHelpTab")!.run!(env)
    expect(dialog.mock.calls).toEqual([["contactSupport"], ["feedback"]])
  })
})

describe("the Help pane", () => {
  beforeEach(() => vi.useFakeTimers())

  function open(mode: "help" | "training" = "help") {
    const c = new SheetController()
    c.setHelpPane(mode)
    const props = { onShowTab: vi.fn(), onDialog: vi.fn(), onWhatsNew: vi.fn(), onDone: vi.fn() }
    const view = render(<HelpPane controller={c} {...props} />)
    return { c, ...props, view }
  }

  it("searches features and shortcuts; Show me opens the tab and points at the button", () => {
    const { onShowTab } = open()
    control("Flash Fill")
    const pane = screen.getByRole("complementary", { name: "Help" })
    fireEvent.change(within(pane).getByRole("searchbox", { name: "Search Help" }), {
      target: { value: "flash" },
    })
    expect(within(pane).getByText("Flash Fill as you type", { selector: "summary" })).toBeTruthy()
    expect(within(pane).queryByText("Forecast Sheet")).toBeNull()
    expect(within(pane).getByText("Ctrl+E")).toBeTruthy()
    fireEvent.click(within(pane).getByRole("button", { name: "Show me" }))
    expect(onShowTab).toHaveBeenCalledWith("TabData")
    act(() => void vi.advanceTimersByTime(200))
    expect(usePointer.getState().pointed).toMatchObject({ id: "flash-fill-live", target: "Flash Fill" })
  })

  it("says so when the feature isn't on screen to point at", () => {
    open()
    const pane = screen.getByRole("complementary", { name: "Help" })
    fireEvent.change(within(pane).getByRole("searchbox"), { target: { value: "quick analysis" } })
    fireEvent.click(within(pane).getByRole("button", { name: "Show me" }))
    act(() => void vi.advanceTimersByTime(200))
    expect(usePointer.getState().pointed).toBeNull()
    expect(within(pane).getByRole("status").textContent).toMatch(/isn't on screen now/)
  })

  it("Show Training walks through each feature, pointing at it, and Done goes back to Help", () => {
    const { c, view } = open("training")
    for (const t of FEATURE_TOPICS) control(t.target)
    act(() => void vi.advanceTimersByTime(200))
    expect(usePointer.getState().pointed).toMatchObject({
      id: FEATURE_TOPICS[0].id,
      step: `1 of ${FEATURE_TOPICS.length}`,
    })
    const pane = screen.getByRole("complementary", { name: "Training" })
    fireEvent.click(within(pane).getByRole("button", { name: /Next/ }))
    act(() => void vi.advanceTimersByTime(200))
    expect(usePointer.getState().pointed?.id).toBe(FEATURE_TOPICS[1].id)
    for (let i = 2; i < FEATURE_TOPICS.length; i++) fireEvent.click(within(pane).getByRole("button", { name: /Next/ }))
    fireEvent.click(within(pane).getByRole("button", { name: "Done" }))
    expect(c.helpPane).toBe("help")
    expect(usePointer.getState().pointed).toBeNull()
    view.unmount()
  })

  it("leads to Show Training, What's New, Contact Support and Feedback", () => {
    const { c, onDialog, onWhatsNew } = open()
    const more = screen.getByRole("region", { name: "More help" })
    fireEvent.click(within(more).getByRole("button", { name: "What's New" }))
    fireEvent.click(within(more).getByRole("button", { name: "Contact Support" }))
    fireEvent.click(within(more).getByRole("button", { name: "Feedback" }))
    fireEvent.click(within(more).getByRole("button", { name: "Show Training" }))
    expect(onWhatsNew).toHaveBeenCalled()
    expect(onDialog.mock.calls).toEqual([["contactSupport"], ["feedback"]])
    expect(c.helpPane).toBe("training")
  })
})

describe("pointing from Help", () => {
  it("shows the callout at once, even with hints turned off; Got it ends it and dismisses the hint", () => {
    vi.useFakeTimers()
    useSettings.setState({ ...DEFAULT_SETTINGS, loaded: true, setupDone: true, hintsEnabled: false })
    control("Forecast Sheet")
    render(<FeatureHints />)
    expect(screen.queryByRole("dialog")).toBeNull()
    act(() =>
      pointAt({
        id: "forecast-sheet",
        title: "Forecast Sheet",
        text: "Carries values forward.",
        target: "Forecast Sheet",
        step: "6 of 8",
      }),
    )
    const callout = screen.getByRole("dialog", { name: "Forecast Sheet" })
    expect(callout.textContent).toContain("6 of 8")
    fireEvent.click(within(callout).getByRole("button", { name: "Got it" }))
    expect(usePointer.getState().pointed).toBeNull()
    expect(screen.queryByRole("dialog")).toBeNull()
    expect(useSettings.getState().hintsDismissed).toEqual(["forecast-sheet"])
    useSettings.setState({ ...DEFAULT_SETTINGS, loaded: false })
  })
})

describe("Feedback and Contact Support", () => {
  it("fill in a GitHub issue, with the version and system only when left in", () => {
    const system = systemLine("0.7.0", "linux", true)
    expect(system).toBe("Cells 0.7.0 on Linux (app)")
    expect(systemLine("0.7.0", "windows", false)).toBe("Cells 0.7.0 in a browser on Windows")
    expect(systemLine("0.7.0", "web", false)).toBe("Cells 0.7.0 in a browser")
    const url = new URL(
      issueUrl({ kind: "idea", title: " Pin sheets ", details: "Keep two tabs at the front.", system }),
    )
    expect(url.origin + url.pathname).toBe("https://github.com/Shani2480/Cells/issues/new")
    expect(url.searchParams.get("title")).toBe("Idea: Pin sheets")
    expect(url.searchParams.get("body")).toContain("Keep two tabs at the front.")
    expect(url.searchParams.get("body")).toContain("Cells 0.7.0 on Linux (app)")
    expect(issueBody({ kind: "like", title: "x", details: "" })).not.toContain("Cells 0.7.0")
    expect(issueBody({ kind: "problem", title: "x", details: "" })).toContain("Steps to make it happen again")
    expect(issueBody({ kind: "idea", title: "x", details: "a".repeat(9000) }).length).toBeLessThan(6100)
  })

  it("opens the issue in the browser only on Open on GitHub, and only with a title", async () => {
    const external = vi.fn(() => Promise.resolve())
    const onClose = vi.fn()
    const platform = { os: "linux", opens: { external } } as unknown as Platform
    render(
      <PlatformContext.Provider value={platform}>
        <FeedbackDialog mode="support" onClose={onClose} />
      </PlatformContext.Provider>,
    )
    const dialog = screen.getByRole("dialog", { name: "Contact Support" })
    const openOnGitHub = within(dialog).getByRole("button", { name: "Open on GitHub" }) as HTMLButtonElement
    expect(openOnGitHub.disabled).toBe(true)
    fireEvent.click(within(dialog).getByRole("radio", { name: "I have a question" }))
    fireEvent.change(within(dialog).getByLabelText("In a few words"), { target: { value: "Where is Goal Seek?" } })
    fireEvent.click(within(dialog).getByRole("checkbox"))
    await act(async () => void fireEvent.click(openOnGitHub))
    expect(external).toHaveBeenCalledTimes(1)
    const url = new URL((external.mock.calls[0] as unknown as [string])[0])
    expect(url.searchParams.get("title")).toBe("Question: Where is Goal Seek?")
    expect(url.searchParams.get("body")).not.toContain("Cells 0.7.0")
    expect(onClose).toHaveBeenCalled()
  })
})
