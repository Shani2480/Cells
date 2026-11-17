import { ChevronDown, ChevronRight, CircleAlert, CircleCheck, TriangleAlert, X } from "lucide-react"
import { useMemo, useState, useSyncExternalStore } from "react"
import {
  checkAccessibility,
  goToIssue,
  ISSUE_INFO,
  ISSUE_ORDER,
  issueActions,
  type Issue,
  type IssueKind,
  type Severity,
} from "@/sheet/accessibility"
import type { SheetController } from "@/sheet/controller"

const SEVERITY_LABEL: Record<Severity, string> = { error: "Errors", warning: "Warnings" }

/**
 * Review › Check Accessibility: Excel's Accessibility pane. Inspection Results lists errors, then warnings, by kind;
 * picking one goes to it and offers the recommended actions, with why and how to fix it below. It checks again
 * after every change.
 */
export function AccessibilityPane({
  controller,
  onDone,
  onDialog,
}: {
  controller: SheetController
  onDone: () => void
  onDialog: (name: string) => void
}) {
  const version = useSyncExternalStore(controller.subscribe, controller.getVersion)
  const issues = useMemo(() => checkAccessibility(controller.wb), [controller, version])
  const [closed, setClosed] = useState<Set<string>>(new Set())
  const [picked, setPicked] = useState<{ kind: IssueKind; index: number } | null>(null)
  const close = () => {
    controller.setAccessibilityPane(false)
    onDone()
  }
  const toggle = (key: string) =>
    setClosed((s) => {
      const next = new Set(s)
      if (next.has(key)) next.delete(key)
      else next.add(key)
      return next
    })
  const byKind = new Map<IssueKind, Issue[]>()
  for (const i of issues) byKind.set(i.kind, [...(byKind.get(i.kind) ?? []), i])
  const current = picked ? byKind.get(picked.kind)?.[picked.index] : undefined
  const info = current ? ISSUE_INFO[current.kind] : null

  return (
    <aside
      className="comments-pane a11y-pane"
      aria-label="Accessibility"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Accessibility</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <div className="comments-pane-list a11y-list">
        <h3 className="a11y-title">Inspection Results</h3>
        {!issues.length ? (
          <p className="a11y-clear">
            <CircleCheck size={18} strokeWidth={1.75} aria-hidden="true" />
            <span>
              No accessibility issues found. People with disabilities shouldn't have difficulty reading this workbook.
            </span>
          </p>
        ) : null}
        {(["error", "warning"] as const).map((severity) => {
          const kinds = ISSUE_ORDER.filter((k) => ISSUE_INFO[k].severity === severity && byKind.has(k))
          if (!kinds.length) return null
          const count = kinds.reduce((n, k) => n + byKind.get(k)!.length, 0)
          const open = !closed.has(severity)
          return (
            <section key={severity} className="a11y-group">
              <button
                type="button"
                className="a11y-row a11y-severity"
                aria-expanded={open}
                onClick={() => toggle(severity)}
              >
                {open ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                {severity === "error" ? (
                  <CircleAlert size={15} className="a11y-icon is-error" aria-hidden="true" />
                ) : (
                  <TriangleAlert size={15} className="a11y-icon is-warning" aria-hidden="true" />
                )}
                <span>
                  {SEVERITY_LABEL[severity]} ({count})
                </span>
              </button>
              {open
                ? kinds.map((kind) => {
                    const list = byKind.get(kind)!
                    const kindOpen = !closed.has(kind)
                    return (
                      <div key={kind} className="a11y-kind">
                        <button
                          type="button"
                          className="a11y-row a11y-kind-head"
                          aria-expanded={kindOpen}
                          onClick={() => toggle(kind)}
                        >
                          {kindOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                          <span>
                            {ISSUE_INFO[kind].title} ({list.length})
                          </span>
                        </button>
                        {kindOpen
                          ? list.map((issue, index) => {
                              const on = picked?.kind === kind && picked.index === index
                              return (
                                <div key={index} className={`a11y-item${on ? " is-picked" : ""}`}>
                                  <button
                                    type="button"
                                    className="a11y-row a11y-item-head"
                                    aria-pressed={on}
                                    onClick={() => {
                                      setPicked({ kind, index })
                                      goToIssue(controller, issue)
                                    }}
                                  >
                                    <span className="a11y-item-label">{issue.label}</span>
                                    <span className="a11y-item-sheet">{controller.wb.sheets[issue.sheet]?.name}</span>
                                  </button>
                                  {on ? (
                                    <div className="a11y-actions" role="group" aria-label="Recommended actions">
                                      {issueActions(controller, issue, onDialog).map((a) => (
                                        <button key={a.label} type="button" className="dialog-btn" onClick={a.run}>
                                          {a.label}
                                        </button>
                                      ))}
                                    </div>
                                  ) : null}
                                </div>
                              )
                            })
                          : null}
                      </div>
                    )
                  })
                : null}
            </section>
          )
        })}
      </div>
      {info ? (
        <footer className="a11y-info">
          <h3>Additional Information</h3>
          <h4>Why Fix:</h4>
          <p>{info.why}</p>
          <h4>How to Fix:</h4>
          <p>{info.how}</p>
        </footer>
      ) : null}
    </aside>
  )
}
