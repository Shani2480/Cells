/**
 * Help › Contact Support and Feedback: a new GitHub issue, filled in for the person to read, change and submit on
 * GitHub themselves. Cells sends nothing; the workbook, file names and settings are never part of it.
 */

export const ISSUES_URL = "https://github.com/Shani2480/Cells/issues/new"

export type FeedbackKind = "like" | "dislike" | "idea" | "problem" | "question"

export const KIND_LABELS: Record<FeedbackKind, string> = {
  like: "I like something",
  dislike: "I don't like something",
  idea: "I have a suggestion",
  problem: "Something isn't working",
  question: "I have a question",
}

const KIND_TITLES: Record<FeedbackKind, string> = {
  like: "Feedback",
  dislike: "Feedback",
  idea: "Idea",
  problem: "Problem",
  question: "Question",
}

/** Kept well under what browsers and GitHub take in an address. */
const MAX_BODY = 6000

export type FeedbackDraft = {
  kind: FeedbackKind
  title: string
  details: string
  /** "Cells 0.7.0 on Linux (app)", when the person leaves it in. */
  system?: string
}

/** "Cells 0.7.0 on Linux (app)", or "… in a browser on Linux" for the web build. */
export function systemLine(version: string, os: string, app: boolean): string {
  const names: Record<string, string> = {
    linux: "Linux",
    macos: "macOS",
    windows: "Windows",
    android: "Android",
  }
  const system = names[os]
  if (app) return `Cells ${version} on ${system ?? os} (app)`
  return system ? `Cells ${version} in a browser on ${system}` : `Cells ${version} in a browser`
}

export function issueBody(draft: FeedbackDraft): string {
  const parts = [`**${KIND_LABELS[draft.kind]}**`, draft.details.trim() || "_(Nothing more written.)_"]
  if (draft.kind === "problem") parts.push("**Steps to make it happen again**\n\n1. \n2. \n\n**What I expected**\n\n")
  if (draft.system) parts.push(`---\n${draft.system}`)
  const body = parts.join("\n\n")
  return body.length > MAX_BODY ? `${body.slice(0, MAX_BODY)}…` : body
}

/** The address of a new issue with the draft filled in. */
export function issueUrl(draft: FeedbackDraft): string {
  const title = `${KIND_TITLES[draft.kind]}: ${draft.title.trim()}`
  const query = new URLSearchParams({ title, body: issueBody(draft) })
  return `${ISSUES_URL}?${query.toString()}`
}
