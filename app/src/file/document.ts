import { Workbook, type Selection } from "@cells/core"
import {
  openText,
  openOds,
  openXls,
  openXlsb,
  openXlsx,
  OdsError,
  XlsbError,
  XlsError,
  saveOds,
  saveText,
  saveXlsx,
  textSaveFor,
  type TextImport,
  type TextSave,
  type XlsxSource,
} from "@cells/io"
import type { FileRef, FileType } from "@cells/platform"
import { SheetController } from "@/sheet/controller"
import { DocumentWindows } from "./windows"
import { startCalculation, type WorkerCalculator } from "@/engine-host/host"
import { previewing } from "@/sheet/quick-analysis"

export const PACKAGE_EXTENSIONS = ["xlsx", "xlsm", "xltx", "xltm"]
export const TEXT_EXTENSIONS = ["csv", "txt", "tsv"]
/** Excel 97-2003 workbooks: Cells opens them as a new .xlsx (it doesn't write .xls), as Excel's Convert does. */
export const OLD_EXTENSIONS = ["xls"]
/** Excel Binary Workbooks: opened as a new .xlsx (.xlsm with macros), as .xls are. */
export const BINARY_EXTENSIONS = ["xlsb"]
/** OpenDocument Spreadsheets (LibreOffice, Google Sheets): opened and saved as they are. */
export const ODS_EXTENSIONS = ["ods"]
/** Everything Open reads. */
export const OPEN_EXTENSIONS = [
  ...PACKAGE_EXTENSIONS,
  ...BINARY_EXTENSIONS,
  ...OLD_EXTENSIONS,
  ...ODS_EXTENSIONS,
  ...TEXT_EXTENSIONS,
]

/** Text files Cells opens: CSV opens at once; .txt and .tsv go through the Text Import Wizard, as in Excel. */
export const TEXT_TYPES: FileType[] = [{ name: "Text Files", extensions: TEXT_EXTENSIONS }]

/** Data › Get Data › From PDF. */
export const PDF_TYPES: FileType[] = [{ name: "PDF Files", extensions: ["pdf"] }]

/** Data › From Picture's file types (HEIC and HEIF: phone photos; Android reads them). */
export const PICTURE_TYPES: FileType[] = [
  { name: "Pictures", extensions: ["png", "jpg", "jpeg", "webp", "bmp", "gif", "heic", "heif"] },
]

/** The Open dialog's types; the first (everything Cells opens) is the default. */
export const OPEN_TYPES: FileType[] = [
  { name: "All Excel Files", extensions: OPEN_EXTENSIONS },
  { name: "Excel Workbook", extensions: ["xlsx"] },
  { name: "Excel Macro-Enabled Workbook", extensions: ["xlsm"] },
  { name: "Excel Binary Workbook", extensions: ["xlsb"] },
  { name: "Excel 97-2003 Workbook", extensions: ["xls"] },
  { name: "Excel Template", extensions: ["xltx", "xltm"] },
  { name: "OpenDocument Spreadsheet", extensions: ["ods"] },
  ...TEXT_TYPES,
]

/** Save As types, Excel's names; the first is the default, the file's own type. */
export function saveTypes(fileName: string): FileType[] {
  const types = [
    { name: "Excel Workbook", extensions: ["xlsx"] },
    { name: "Excel Macro-Enabled Workbook", extensions: ["xlsm"] },
    { name: "Excel Template", extensions: ["xltx"] },
    { name: "CSV UTF-8 (Comma delimited)", extensions: ["csv"] },
    { name: "Text (Tab delimited)", extensions: ["txt"] },
    { name: "OpenDocument Spreadsheet", extensions: ["ods"] },
  ]
  const own = types.findIndex((t) => t.extensions.includes(extensionOf(fileName)))
  return own > 0 ? [types[own], ...types.filter((_, i) => i !== own)] : types
}

export const extensionOf = (name: string) => /\.([^.]+)$/.exec(name)?.[1]?.toLowerCase() ?? ""
export const isTextFile = (name: string) => TEXT_EXTENSIONS.includes(extensionOf(name))
export const isOdsFile = (name: string) => ODS_EXTENSIONS.includes(extensionOf(name))
/** The name an .xlsx copy of the workbook takes (AutoRecover keeps one even for a CSV). */
export const packageName = (fileName: string) =>
  PACKAGE_EXTENSIONS.includes(extensionOf(fileName)) ? fileName : `${titleOf(fileName)}.xlsx`
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
  /** A text file's own way of writing (delimiter, encoding), so Save writes it back the same; null for a workbook. */
  text: TextSave | null
  /** Edits (operations applied) since the workbook was made; compared with the count at the last save. */
  edits = 0
  savedEdits = 0
  /** Edits counted at the last AutoRecover copy. */
  recoveredEdits = 0
  /**
   * Inkspire's Edit Data: the name of the document whose chart this workbook holds ("Report.docx"). Such a workbook
   * is a copy Inkspire made and reads back, so each change is saved to it without asking, as Excel does with a chart's
   * workbook from Word.
   */
  embeddedFor: string | null = null
  /** Said yes to "Some features might be lost" for OpenDocument, so Save doesn't ask again. */
  odsAccepted = false
  readonly warnings: string[]
  /** View › New Window's windows; `controller` is the active one. */
  readonly windows: DocumentWindows
  private readonly unsubscribe: () => void

  constructor(
    main: SheetController,
    o: {
      ref?: FileRef | null
      fileName?: string
      source?: XlsxSource | null
      text?: TextSave | null
      warnings?: string[]
    } = {},
    private readonly calc: WorkerCalculator | null = null,
  ) {
    this.ref = o.ref ?? null
    this.fileName = o.fileName ?? "Book1.xlsx"
    this.source = o.source ?? null
    this.text = o.text ?? null
    this.warnings = o.warnings ?? []
    this.unsubscribe = main.wb.subscribe((_changes, ops) => {
      // A Quick Analysis preview comes and goes without changing the workbook.
      if (ops.length && !previewing.has(main.wb)) {
        this.edits++
        this.controller.changed()
      }
    })
    // Changes outside undo (View › Custom Views) count as unsaved too, as in Excel.
    const adopt = (c: SheetController) => {
      c.unsavedChange = () => {
        this.edits++
        c.changed()
      }
      c.printFile = main.printFile
    }
    adopt(main)
    this.windows = new DocumentWindows(main, adopt)
  }

  /** The active window's controller (the workbook's only one unless View › New Window made more). */
  get controller(): SheetController {
    return this.windows.active
  }

  get title() {
    return this.embeddedFor ? `Chart in ${this.embeddedFor}` : titleOf(this.fileName)
  }

  /** The title bar's name: "Budget", or "Budget:2" for the second of the workbook's windows. */
  get windowTitle() {
    return this.windows.title(this.title)
  }

  /** Changed since it was opened or last saved (a new, untouched workbook isn't). */
  get dirty() {
    return this.edits !== this.savedEdits
  }

  /** The file's bytes in the format its name asks for: a workbook, or the active sheet as CSV or text. */
  async bytes(fileName = this.fileName): Promise<Uint8Array> {
    const c = this.controller
    if (c.edit) c.commitEdit()
    if (isTextFile(fileName)) return saveText(c.wb, c.sheetIndex, this.textSave(fileName))
    if (isOdsFile(fileName)) return saveOds(c.wb, { views: c.saveViews(), activeSheet: c.sheetIndex })
    // Every window's view, in order (View › New Window): Excel opens them all again.
    const [first, ...others] = this.windows.all
    return saveXlsx(c.wb, {
      source: this.source,
      fileName,
      views: first.saveViews(),
      activeSheet: first.sheetIndex,
      windows: others.map((w) => ({ views: w.saveViews(), activeSheet: w.sheetIndex })),
    })
  }

  /** How a text file of this name is written: the way it was read when it's the same kind of file. */
  private textSave(fileName: string): TextSave {
    return this.text && extensionOf(fileName) === extensionOf(this.fileName) ? this.text : textSaveFor(fileName)
  }

  /** After a successful save to `ref`. */
  saved(ref: FileRef, edits: number) {
    // Saved somewhere else (Save As): an ordinary workbook from now on.
    if (this.embeddedFor && ref.path !== this.ref?.path) this.embeddedFor = null
    this.text = isTextFile(ref.name) ? this.textSave(ref.name) : null
    this.ref = ref
    this.fileName = ref.name
    this.savedEdits = edits
    this.recoveredEdits = edits
  }

  close() {
    this.unsubscribe()
    this.windows.dispose()
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
  /** For a text file: the Text Import Wizard's choices (a .csv opens with the defaults). */
  text?: Partial<TextImport>,
): Promise<WorkbookDocument> {
  const ext = extensionOf(ref.name)
  if (!OPEN_EXTENSIONS.includes(ext)) throw new FileProblem("type", ref.name)
  let last = 0
  const options = {
    calculate: !worker,
    progress: async (done: number) => {
      progress?.(done)
      // Yield about every 50 ms of work.
      const now = performance.now()
      if (now - last > 50) {
        last = now
        await nextFrame()
      }
    },
  }
  if (TEXT_EXTENSIONS.includes(ext)) {
    const read = await openText(bytes, ref.name, { ...text, ...options })
    const wb = read.workbook as unknown as Workbook<Selection>
    const controller = new SheetController(wb)
    controller.loadViews(read.views, read.activeSheet)
    const calc = worker ? startCalculation(wb, { bytes, name: ref.name, text: text ?? {} }) : null
    const save: TextSave = { delimiter: read.delimiter, encoding: read.encoding, bom: read.bom }
    return new WorkbookDocument(controller, { ref, fileName: ref.name, text: save, warnings: read.warnings }, calc)
  }
  if (OLD_EXTENSIONS.includes(ext)) {
    const read = await openXls(bytes, ref.name, options)
    const wb = read.workbook as unknown as Workbook<Selection>
    const controller = new SheetController(wb)
    controller.loadViews(read.views, read.activeSheet)
    const calc = worker ? startCalculation(wb, { bytes, name: ref.name }) : null
    // A converted copy, not yet saved: Save asks where to keep the .xlsx, and the .xls stays as it was.
    const warnings = [
      `${ref.name} is an Excel 97-2003 workbook. Cells opened a copy; saving it makes an .xlsx and leaves the .xls as it was.`,
      ...read.warnings,
    ]
    return new WorkbookDocument(controller, { ref: null, fileName: `${titleOf(ref.name)}.xlsx`, warnings }, calc)
  }
  if (BINARY_EXTENSIONS.includes(ext)) {
    const read = await openXlsb(bytes, ref.name, options)
    const wb = read.workbook as unknown as Workbook<Selection>
    const controller = new SheetController(wb)
    controller.loadViews(read.views, read.activeSheet)
    const calc = worker ? startCalculation(wb, { bytes, name: ref.name }) : null
    // A converted copy, as with an .xls: Save makes an .xlsx (an .xlsm when it has macros) and leaves the .xlsb be.
    const kind = read.macros ? "xlsm" : "xlsx"
    const warnings = [
      `${ref.name} is an Excel Binary Workbook. Cells opened a copy; saving it makes an .${kind} and leaves the .xlsb as it was.`,
      ...read.warnings,
    ]
    return new WorkbookDocument(
      controller,
      { ref: null, fileName: `${titleOf(ref.name)}.${kind}`, source: read.source, warnings },
      calc,
    )
  }
  if (ODS_EXTENSIONS.includes(ext)) {
    const read = await openOds(bytes, ref.name, options)
    const wb = read.workbook as unknown as Workbook<Selection>
    const controller = new SheetController(wb)
    controller.loadViews(read.views, read.activeSheet)
    const calc = worker ? startCalculation(wb, { bytes, name: ref.name }) : null
    return new WorkbookDocument(controller, { ref, fileName: ref.name, warnings: read.warnings }, calc)
  }
  const read = await openXlsx(bytes, ref.name, options)
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
  // The other windows the file was saved with, each with its own sheet, scroll, zoom and selection; the first stays
  // in front.
  for (const w of read.windows) doc.windows.newWindow().loadViews(w.views, w.activeSheet)
  if (read.windows.length) doc.windows.activate(controller)
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
      return `'${name}' is in the Excel 5.0/95 format or older, which Cells can't open. Save it as .xlsx in Excel first.`
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
  if (error instanceof XlsError || error instanceof XlsbError || error instanceof OdsError)
    return new FileProblem(error.kind, name, detail)
  // A compound (OLE) file named .xlsx: encrypted (an .xls opens through its own reader).
  if (bytes && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0)
    return new FileProblem(extensionOf(name) === "xls" ? "damaged" : "encrypted", name, detail)
  if (!bytes) return new FileProblem("read", name, detail)
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return new FileProblem("type", name, detail)
  return new FileProblem("damaged", name, detail)
}
