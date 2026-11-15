import { MARGIN_PRESETS, PAPER_SIZES, paperOf } from "@cells/core"
import { ChevronLeft, ChevronRight, FileText, Printer } from "lucide-react"
import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react"
import type { FileCommands } from "@/file/commands"
import { jobPdf, printJob, printJobPages, type JobPage, type PrintWhat } from "@/print/job"
import { drawPage, type PrintContext } from "@/print/render"
import type { SheetController } from "@/sheet/controller"
import { marginPreset } from "@/sheet/print"
import { marginUnit, type PageSetupTab } from "./PageSetupDialog"

const WHAT: { id: PrintWhat; label: string; hint: string }[] = [
  { id: "active", label: "Print Active Sheets", hint: "Only print the active sheets" },
  { id: "workbook", label: "Print Entire Workbook", hint: "Print the entire workbook" },
  { id: "selection", label: "Print Selection", hint: "Only print the current selection" },
]

type Scaling = "none" | "sheet" | "columns" | "rows" | "custom"
const SCALING: { id: Scaling; label: string; hint: string }[] = [
  { id: "none", label: "No Scaling", hint: "Print sheets at their actual size" },
  { id: "sheet", label: "Fit Sheet on One Page", hint: "Shrink the printout so that it fits on one page" },
  { id: "columns", label: "Fit All Columns on One Page", hint: "Shrink the printout so that it is one page wide" },
  { id: "rows", label: "Fit All Rows on One Page", hint: "Shrink the printout so that it is one page high" },
]

function scalingOf(c: SheetController): Scaling {
  const s = c.print.setup()
  const f = s.fitTo
  if (!f) return s.scale === 100 ? "none" : "custom"
  if (f.width === 1 && f.height === 1) return "sheet"
  if (f.width === 1 && f.height === 0) return "columns"
  if (f.width === 0 && f.height === 1) return "rows"
  return "custom"
}

export type PrintFile = { fileName: string; folder: string }

/** What the header and footer codes print, for this workbook now. */
export function printContext(file: PrintFile, total: number): PrintContext {
  return { fileName: file.fileName, folder: file.folder, now: new Date(), total }
}

/**
 * File › Print, as Excel's: the settings on the left (what prints, pages, orientation, paper, margins, scaling) and the
 * pages as they will print on the right. Print hands them to the system's print dialog.
 */
export function PrintPage({
  controller: c,
  file,
  onPageSetup,
  onDone,
}: {
  controller: SheetController
  file: PrintFile
  onPageSetup: (tab: PageSetupTab) => void
  onDone: () => void
}) {
  const version = useSyncExternalStore(c.subscribe, c.getVersion)
  const [what, setWhat] = useState<PrintWhat>("active")
  const [ignoreArea, setIgnoreArea] = useState(false)
  const [from, setFrom] = useState("")
  const [to, setTo] = useState("")
  const [index, setIndex] = useState(0)
  const [busy, setBusy] = useState(false)
  // Once previewed, the sheet shows where its pages break.
  useEffect(() => {
    c.print.showBreaks = true
  }, [c])

  // version: the workbook changed, so its pages may have.
  const all = useMemo(() => printJob(c, what, ignoreArea), [c, what, ignoreArea, version])
  const pages = useMemo(() => {
    const a = Number(from) || 1
    const b = Number(to) || all.length
    return all.filter((_, i) => i + 1 >= a && i + 1 <= b)
  }, [all, from, to])
  const shown = Math.min(index, Math.max(0, pages.length - 1))
  const setup = c.print.setup()
  const unit = marginUnit(c)
  const inches = (v: number) => `${Math.round(v * unit.perInch * 100) / 100}${unit.name === "in" ? '"' : " cm"}`
  const paper = PAPER_SIZES[paperOf(setup, c.print.defaultPaper)]
  const preset = marginPreset(setup.margins)
  const scaling = scalingOf(c)

  const print = async () => {
    setBusy(true)
    try {
      await printJobPages(c, pages, printContext(file, all.length))
      onDone()
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="backstage-content print-page">
      <div className="print-settings">
        <h1>Print</h1>
        <button type="button" className="print-go" onClick={() => void print()} disabled={!pages.length || busy}>
          <Printer size={28} strokeWidth={1.5} />
          <span>Print</span>
        </button>
        <h2 className="print-heading">Printer</h2>
        <p className="print-hint">You pick the printer, and the number of copies, in the print dialog.</p>
        <h2 className="print-heading">Settings</h2>
        <Setting
          label="What to print"
          value={what}
          options={WHAT.map((w) => ({ value: w.id, label: w.label, hint: w.hint }))}
          onChange={(v) => {
            setWhat(v as PrintWhat)
            setIndex(0)
          }}
        />
        <label className="print-check">
          <input type="checkbox" checked={ignoreArea} onChange={(e) => setIgnoreArea(e.target.checked)} />
          Ignore Print Area
        </label>
        <div className="print-pages">
          <span>Pages:</span>
          <input
            aria-label="From page"
            className="dialog-field"
            inputMode="numeric"
            value={from}
            onChange={(e) => setFrom(e.target.value.replace(/\D/g, ""))}
          />
          <span>to</span>
          <input
            aria-label="To page"
            className="dialog-field"
            inputMode="numeric"
            value={to}
            onChange={(e) => setTo(e.target.value.replace(/\D/g, ""))}
          />
        </div>
        <Setting
          label="Orientation"
          value={setup.orientation}
          options={[
            { value: "portrait", label: "Portrait Orientation" },
            { value: "landscape", label: "Landscape Orientation" },
          ]}
          onChange={(v) => c.print.setOrientation(v as "portrait" | "landscape")}
        />
        <Setting
          label="Paper size"
          value={String(paperOf(setup, c.print.defaultPaper))}
          options={Object.entries(PAPER_SIZES).map(([id, p]) => ({
            value: id,
            label: p.name,
            hint: `${inches(p.width)} x ${inches(p.height)}`,
          }))}
          hint={paper ? `${inches(paper.width)} x ${inches(paper.height)}` : undefined}
          onChange={(v) => c.print.setPaper(Number(v))}
        />
        <Setting
          label="Margins"
          value={preset}
          options={[
            ...(preset === "custom" ? [{ value: "custom", label: "Custom Margins" }] : []),
            ...(Object.keys(MARGIN_PRESETS) as (keyof typeof MARGIN_PRESETS)[]).map((k) => ({
              value: k,
              label: `${k[0].toUpperCase()}${k.slice(1)} Margins`,
            })),
            { value: "dialog", label: "Custom Margins..." },
          ]}
          hint={`Left: ${inches(setup.margins.left)}   Right: ${inches(setup.margins.right)}`}
          onChange={(v) => {
            if (v === "dialog") onPageSetup("margins")
            else if (v in MARGIN_PRESETS) c.print.setMargins(MARGIN_PRESETS[v as keyof typeof MARGIN_PRESETS])
          }}
        />
        <Setting
          label="Scaling"
          value={scaling}
          options={[
            ...SCALING.map((s) => ({ value: s.id, label: s.label, hint: s.hint })),
            ...(scaling === "custom" ? [{ value: "custom", label: "Custom Scaling" }] : []),
            { value: "dialog", label: "Custom Scaling Options..." },
          ]}
          onChange={(v) => {
            if (v === "dialog") onPageSetup("page")
            else
              c.print.change((s) => {
                s.fitTo =
                  v === "sheet"
                    ? { width: 1, height: 1 }
                    : v === "columns"
                      ? { width: 1, height: 0 }
                      : v === "rows"
                        ? { width: 0, height: 1 }
                        : null
                if (v === "none") s.scale = 100
              })
          }}
        />
        <button type="button" className="print-link" onClick={() => onPageSetup("page")}>
          Page Setup
        </button>
      </div>
      <div className="print-preview">
        <PagePreview controller={c} job={pages[shown] ?? null} ctx={printContext(file, all.length)} version={version} />
        <div className="print-nav">
          <button type="button" aria-label="Previous Page" disabled={shown === 0} onClick={() => setIndex(shown - 1)}>
            <ChevronLeft size={16} />
          </button>
          <input
            aria-label="Current Page"
            className="dialog-field"
            value={pages.length ? shown + 1 : 0}
            onChange={(e) => {
              const n = Number(e.target.value)
              if (n >= 1 && n <= pages.length) setIndex(n - 1)
            }}
          />
          <span>of {pages.length}</span>
          <button
            type="button"
            aria-label="Next Page"
            disabled={shown >= pages.length - 1}
            onClick={() => setIndex(shown + 1)}
          >
            <ChevronRight size={16} />
          </button>
        </div>
      </div>
    </div>
  )
}

/** One of Print's settings: Excel shows a wide button with the choice and a line under it; here a list does. */
function Setting({
  label,
  value,
  options,
  hint,
  onChange,
}: {
  label: string
  value: string
  options: { value: string; label: string; hint?: string }[]
  hint?: string
  onChange: (value: string) => void
}) {
  const current = options.find((o) => o.value === value)
  return (
    <label className="print-setting">
      <select aria-label={label} value={value} onChange={(e) => onChange(e.target.value)}>
        {options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
      <span className="print-setting-face" aria-hidden>
        <strong>{current?.label ?? ""}</strong>
        <span>{hint ?? current?.hint ?? ""}</span>
      </span>
    </label>
  )
}

/** The page as it will print, fitted to the preview area (drawn again when the workbook changes). */
function PagePreview({
  controller,
  job,
  ctx,
  version,
}: {
  controller: SheetController
  job: JobPage | null
  ctx: PrintContext
  version: number
}) {
  const box = useRef<HTMLDivElement>(null)
  const canvas = useRef<HTMLCanvasElement>(null)
  const [size, setSize] = useState({ w: 0, h: 0 })
  useEffect(() => {
    const el = box.current
    if (!el) return
    const observer = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }))
    observer.observe(el)
    return () => observer.disconnect()
  }, [])
  const { fileName, folder, total } = ctx
  useEffect(() => {
    const cv = canvas.current
    if (!cv || !job || !size.w) return
    let live = true
    const { width, height } = job.pagination.paper
    const fit = Math.min((size.w - 24) / width, (size.h - 24) / height)
    const dpr = window.devicePixelRatio || 1
    cv.style.width = `${Math.floor(width * fit)}px`
    cv.style.height = `${Math.floor(height * fit)}px`
    const off = document.createElement("canvas")
    off.width = Math.round(width * fit * dpr)
    off.height = Math.round(height * fit * dpr)
    void drawPage(
      off,
      controller,
      job.sheet,
      job.pagination,
      job.page,
      { fileName, folder, now: new Date(), total },
      fit * dpr,
    ).then(() => {
      if (!live) return
      cv.width = off.width
      cv.height = off.height
      cv.getContext("2d")!.drawImage(off, 0, 0)
    })
    return () => {
      live = false
    }
  }, [controller, job, size, version, fileName, folder, total])
  return (
    <div className="print-sheet-area" ref={box}>
      {job ? (
        <canvas ref={canvas} className="print-paper" data-keep-light aria-label="Print Preview" />
      ) : (
        <p className="print-empty">We didn't find anything to print.</p>
      )}
    </div>
  )
}

/** File › Export: Create PDF/XPS Document (PDF here). */
export function ExportPage({
  controller: c,
  file,
  files,
  run,
}: {
  controller: SheetController
  file: PrintFile
  files: FileCommands
  run: (command: () => Promise<boolean>) => void
}) {
  const [what, setWhat] = useState<PrintWhat>("active")
  const create = () =>
    run(() =>
      files.exportPdf((progress) => {
        const pages = printJob(c, what)
        return jobPdf(c, pages, printContext(file, pages.length), progress)
      }),
    )
  return (
    <div className="backstage-content">
      <h1>Export</h1>
      <div className="bs-split">
        <div className="bs-places">
          <button type="button" className="bs-place is-current">
            <FileText size={20} strokeWidth={1.5} />
            Create PDF/XPS Document
          </button>
        </div>
        <div className="bs-split-main">
          <h2 className="print-heading">Create a PDF/XPS Document</h2>
          <ul className="bs-bullets">
            <li>Preserves layout, formatting, fonts, and images</li>
            <li>Content can't be easily changed</li>
            <li>Free viewers are available on the web</li>
          </ul>
          <label className="print-setting export-what">
            <select aria-label="Publish what" value={what} onChange={(e) => setWhat(e.target.value as PrintWhat)}>
              {WHAT.map((w) => (
                <option key={w.id} value={w.id}>
                  {w.label.replace("Print ", "")}
                </option>
              ))}
            </select>
            <span className="print-setting-face" aria-hidden>
              <strong>{WHAT.find((w) => w.id === what)!.label.replace("Print ", "")}</strong>
              <span>{WHAT.find((w) => w.id === what)!.hint.replace("print", "publish")}</span>
            </span>
          </label>
          <button type="button" className="bs-button" onClick={create}>
            Create PDF
          </button>
        </div>
      </div>
    </div>
  )
}
