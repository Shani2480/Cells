import { rangeName } from "@cells/core"
import type { LinkPort } from "@cells/platform"
import type { SheetController } from "@/sheet/controller"
import { a1Range } from "./link"

/**
 * Live links for Inkspire's linked tables (docs/linked-ranges.md). Inkspire lists the ranges its open documents
 * link to in `links/wanted.json`; while the workbook is open here, Cells writes what those ranges hold to
 * `links/<snapshotName(path)>.json` a moment after they change (an edit, a recalculation, a save), so Inkspire can
 * update the linked table without the file being saved. Only a changed snapshot is written.
 */

export type LinkWant = { file: string; range: string }

export type RangeSnapshot =
  | {
      sheet: string
      /** "A1:D10". */
      range: string
      text: string
      html: string
      /** Each cell's value: a number, text, true/false, an error code ("#REF!") or null. */
      values: (number | string | boolean | null)[][]
    }
  | { error: string }

export type LinkSnapshot = {
  app: "cells"
  version: 1
  file: string
  /** When Cells wrote this (ISO time). */
  updated: string
  /** False while the workbook has changes not saved to `file`. */
  saved: boolean
  /** Keyed by the range exactly as `wanted.json` gives it. */
  ranges: Record<string, RangeSnapshot>
}

/** A linked range bigger than this isn't written (a table that size doesn't belong in a document). */
export const MAX_LINKED_CELLS = 20_000

/** The ranges Inkspire wants, from `wanted.json`'s text; bad entries are skipped. */
export function parseWants(text: string | null): LinkWant[] {
  if (!text) return []
  try {
    const data = JSON.parse(text) as { links?: unknown }
    if (!Array.isArray(data.links)) return []
    return data.links.filter(
      (l): l is LinkWant =>
        !!l && typeof (l as LinkWant).file === "string" && typeof (l as LinkWant).range === "string",
    )
  } catch {
    return []
  }
}

/** Windows paths are the same file whatever their case and slashes. */
export function samePath(a: string, b: string): boolean {
  const windows = (p: string) => /^[A-Za-z]:[\\/]/.test(p) || p.startsWith("\\\\")
  if (windows(a) && windows(b)) return a.replace(/\//g, "\\").toLowerCase() === b.replace(/\//g, "\\").toLowerCase()
  return a === b
}

/** The snapshot's file name for a workbook path: FNV-1a (64-bit) of its UTF-8 bytes, as 16 hex digits. */
export function snapshotName(path: string): string {
  let hash = 0xcbf29ce484222325n
  for (const byte of new TextEncoder().encode(path)) {
    hash ^= BigInt(byte)
    hash = (hash * 0x100000001b3n) & 0xffffffffffffffffn
  }
  return hash.toString(16).padStart(16, "0")
}

/** What the wanted ranges of this workbook hold now. */
export function snapshotRanges(controller: SheetController, wants: readonly string[]): Record<string, RangeSnapshot> {
  const ranges: Record<string, RangeSnapshot> = {}
  for (const want of wants) {
    const target = controller.names.resolve(a1Range(want))
    if (!target) {
      ranges[want] = { error: "#REF!" }
      continue
    }
    const { range } = target
    if ((range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1) > MAX_LINKED_CELLS) {
      ranges[want] = { error: "too big" }
      continue
    }
    const content = controller.rangeContent(target.sheet, range)
    ranges[want] = {
      sheet: controller.wb.sheets[target.sheet].name,
      range: rangeName(range),
      text: content.text,
      html: content.html,
      values: content.values,
    }
  }
  return ranges
}

export type LiveDocument = { path: string | null; dirty: boolean; controller: SheetController }

/** Writes snapshots for the shown workbook; `schedule` after anything that may change them. */
export class LiveLinks {
  private timer: ReturnType<typeof setTimeout> | null = null
  /** What was last written per snapshot name (without its time), so an unchanged one isn't written again. */
  private readonly written = new Map<string, string>()
  private running: Promise<boolean> | null = null

  constructor(
    private readonly port: LinkPort,
    private readonly current: () => LiveDocument,
    private readonly delay = 400,
  ) {}

  schedule() {
    if (this.timer) clearTimeout(this.timer)
    this.timer = setTimeout(() => {
      this.timer = null
      void this.refresh()
    }, this.delay)
  }

  stop() {
    if (this.timer) clearTimeout(this.timer)
    this.timer = null
  }

  /** Writes the shown workbook's snapshot if Inkspire wants any of its ranges and it changed. */
  async refresh(): Promise<boolean> {
    // One at a time: a slow write must not let an older snapshot land after a newer one.
    while (this.running) await this.running
    this.running = this.write().finally(() => (this.running = null))
    return this.running
  }

  private async write(): Promise<boolean> {
    const doc = this.current()
    const path = doc.path
    if (!path) return false
    let wanted: string | null
    try {
      wanted = await this.port.wanted()
    } catch {
      return false
    }
    const wants = [
      ...new Set(
        parseWants(wanted)
          .filter((w) => samePath(w.file, path))
          .map((w) => w.range),
      ),
    ]
    if (!wants.length) return false
    const body = { app: "cells" as const, version: 1 as const, file: path, saved: !doc.dirty }
    const ranges = snapshotRanges(doc.controller, wants)
    const key = JSON.stringify({ ...body, ranges })
    const name = snapshotName(path)
    if (this.written.get(name) === key) return false
    const snapshot: LinkSnapshot = { ...body, updated: new Date().toISOString(), ranges }
    try {
      await this.port.publish(name, JSON.stringify(snapshot))
    } catch {
      return false
    }
    this.written.set(name, key)
    return true
  }
}
