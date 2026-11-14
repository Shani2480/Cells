import { Workbook, type Selection } from "@cells/core"
import { openXlsx, saveXlsx, type XlsxSource } from "@cells/io"
import type { FileRef, FileType } from "@cells/platform"
import { SheetController } from "@/sheet/controller"
import { startCalculation, type WorkerCalculator } from "@/engine-host/host"

/** The workbook formats Cells opens and saves (CSV follows in this phase). */
export const OPEN_TYPES: FileType[] = [
  { name: "All Excel Files", extensions: ["xlsx", "xlsm", "xltx", "xltm"] },
  { name: "Excel Workbook", extensions: ["xlsx"] },
  { name: "Excel Macro-Enabled Workbook", extensions: ["xlsm"] },
  { name: "Excel Template", extensions: ["xltx", "xltm"] },
]

/** Save As types, Excel's names; the first is the default. A macro-enabled file keeps its type first. */
export function saveTypes(fileName: string): FileType[] {
  const workbook = { name: "Excel Workbook", extensions: ["xlsx"] }
  const macro = { name: "Excel Macro-Enabled Workbook", extensions: ["xlsm"] }
  const template = { name: "Excel Template", extensions: ["xltx"] }
  return /\.xlsm$/i.test(fileName) ? [macro, workbook, template] : [workbook, macro, template]
}

export const extensionOf = (name: string) => /\.([^.]+)$/.exec(name)?.[1]?.toLowerCase() ?? ""
/** The title bar's name: the file name without its extension. */
export const titleOf = (fileName: string) => fileName.replace(/\.[^.]+$/, "")

let documents = 0

/**
 * One open workbook: its controller, where it lives (null until first saved), what was read from the file (so a save
 * keeps what Cells doesn't model), and whether it changed since it was last saved.
 */
export class WorkbookDocument {
  readonly id = `doc-${Date.now().toString(36)}-${++documents}`
  ref: FileRef | null
  /** "Budget.xlsx", or "Book1.xlsx" for a new workbook. */
  fileName: string
  source: XlsxSource | null
  /** Edits (operations applied) since the workbook was made; compared with the count at the last save. */
  edits = 0
  savedEdits = 0
  /** Edits counted at the last AutoRecover copy. */
  recoveredEdits = 0
  readonly warnings: string[]
  private readonly unsubscribe: () => void

  constructor(
    readonly controller: SheetController,
    o: { ref?: FileRef | null; fileName?: string; source?: XlsxSource | null; warnings?: string[] } = {},
    private readonly calc: WorkerCalculator | null = null,
  ) {
    this.ref = o.ref ?? null
    this.fileName = o.fileName ?? "Book1.xlsx"
    this.source = o.source ?? null
    this.warnings = o.warnings ?? []
    this.unsubscribe = controller.wb.subscribe((_changes, ops) => {
      if (ops.length) {
        this.edits++
        this.controller.changed()
      }
    })
  }

  get title() {
    return titleOf(this.fileName)
  }

  /** Changed since it was opened or last saved (a new, untouched workbook isn't). */
  get dirty() {
    return this.edits !== this.savedEdits
  }

  /** The bytes of an .xlsx save (for Save, Save As and AutoRecover). */
  async bytes(fileName = this.fileName): Promise<Uint8Array> {
    const c = this.controller
    if (c.edit) c.commitEdit()
    return saveXlsx(c.wb, {
      source: this.source,
      fileName,
      views: c.saveViews(),
      activeSheet: c.sheetIndex,
    })
  }

  /** After a successful save to `ref`. */
  saved(ref: FileRef, edits: number) {
    this.ref = ref
    this.fileName = ref.name
    this.savedEdits = edits
    this.recoveredEdits = edits
  }

  close() {
    this.unsubscribe()
    this.calc?.stop()
  }
}

/** A new, empty workbook ("Book1.xlsx" unless named), calculating in a worker. */
export function newDocument(controller?: SheetController, fileName = "Book1.xlsx"): WorkbookDocument {
  const c = controller ?? new SheetController()
  const calc = controller ? null : startCalculation(c.wb)
  return new WorkbookDocument(c, { fileName }, calc)
}

/** Waits for the next frame (or a moment when there are no frames, as in tests), so the window can draw. */
export const nextFrame = () =>
  new Promise<void>((resolve) => {
    if (typeof requestAnimationFrame === "function" && typeof document !== "undefined" && !document.hidden)
      requestAnimationFrame(() => resolve())
    else setTimeout(resolve, 0)
  })

/**
 * Opens a workbook from its bytes. Reading yields to the window between slices, so a big file shows its progress
 * instead of freezing the app; the calculation worker opens the same bytes on its own.
 */
export async function openDocument(
  bytes: Uint8Array,
  ref: FileRef,
  progress?: (done: number) => void,
  worker = true,
): Promise<WorkbookDocument> {
  const ext = extensionOf(ref.name)
  if (!["xlsx", "xlsm", "xltx", "xltm"].includes(ext)) throw new FileProblem("type", ref.name)
  let last = 0
  const read = await openXlsx(bytes, ref.name, {
    calculate: !worker,
    progress: async (done) => {
      progress?.(done)
      // Yield about every 50 ms of work.
      const now = performance.now()
      if (now - last > 50) {
        last = now
        await nextFrame()
      }
    },
  })
  const wb = read.workbook as unknown as Workbook<Selection>
  const controller = new SheetController(wb)
  controller.loadViews(read.views, read.activeSheet)
  const calc = worker ? startCalculation(wb, { bytes, name: ref.name }) : null
  // A template opens as a new workbook based on it (Budget1), as Excel does; saving asks where.
  const template = ext === "xltx" || ext === "xltm"
  const doc = new WorkbookDocument(
    controller,
    {
      ref: template ? null : ref,
      fileName: template ? `${titleOf(ref.name)}1.${ext === "xltm" ? "xlsm" : "xlsx"}` : ref.name,
      source: read.source,
      warnings: read.warnings,
    },
    calc,
  )
  return doc
}

/** Why a file couldn't be opened or saved, in words for the person (the raw error goes under Details). */
export class FileProblem extends Error {
  constructor(
    readonly kind: "type" | "damaged" | "encrypted" | "old" | "read" | "write",
    readonly fileName: string,
    readonly detail = "",
  ) {
    super(problemMessage(kind, fileName))
    this.name = "FileProblem"
  }
}

export function problemMessage(kind: FileProblem["kind"], name: string): string {
  switch (kind) {
    case "type":
      return `Cells can't open '${name}' because the file format or file extension is not valid. Verify that the file has not been corrupted and that the file extension matches the format of the file.`
    case "damaged":
      return `We found a problem with some content in '${name}'. The file is damaged and can't be opened.`
    case "encrypted":
      return `'${name}' is protected with a password to open. Cells can't open password-protected files yet.`
    case "old":
      return `'${name}' is in the Excel 97-2003 format (.xls), which Cells can't open yet. Save it as .xlsx in Excel first.`
    case "read":
      return `Sorry, we couldn't open '${name}'. It may have been moved, renamed or deleted.`
    case "write":
      return `Cells couldn't save '${name}'. Check that the folder still exists and that the file isn't open elsewhere or read-only.`
  }
}

/** Sorts a reading error into a FileProblem by what the file's first bytes are. */
export function classifyOpenError(error: unknown, bytes: Uint8Array | null, name: string): FileProblem {
  if (error instanceof FileProblem) return error
  const detail = error instanceof Error ? error.message : String(error)
  // A compound (OLE) file: an encrypted .xlsx or an old .xls.
  if (bytes && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0)
    return new FileProblem(extensionOf(name) === "xls" ? "old" : "encrypted", name, detail)
  if (!bytes) return new FileProblem("read", name, detail)
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return new FileProblem("type", name, detail)
  return new FileProblem("damaged", name, detail)
}
