import {
  columnName,
  copyPageSetup,
  PAPER_SIZES,
  paperOf,
  parseRange,
  type HeaderFooterText,
  type PageSetup,
  type PrintPage,
  type Range,
  type Sheet,
} from "@cells/core"
import { headerFooterLines, type PrintContext } from "@/print/render"
import { useState, type ReactNode } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

export type PageSetupTab = "page" | "margins" | "headerFooter" | "sheet"

const TABS: { id: PageSetupTab; label: string }[] = [
  { id: "page", label: "Page" },
  { id: "margins", label: "Margins" },
  { id: "headerFooter", label: "Header/Footer" },
  { id: "sheet", label: "Sheet" },
]

/** Excel's message for a print area or title it can't read. */
export const BAD_REFERENCE = "The text you entered is not a valid reference or defined name."

/** Margins show in the region's unit: inches where Letter is the paper, centimetres elsewhere. */
export function marginUnit(c: SheetController): { name: string; perInch: number; step: number } {
  return c.print.defaultPaper === 1 ? { name: "in", perInch: 1, step: 0.25 } : { name: "cm", perInch: 2.54, step: 0.5 }
}

/** The Header and Footer lists: Excel's presets that need no author name (from memory of Excel 365). */
export function headerPresets(sheetName: string, fileName: string): { label: string; value: HeaderFooterText }[] {
  const hf = (left: string, center: string, right: string) => ({ left, center, right })
  return [
    { label: "(none)", value: hf("", "", "") },
    { label: "Page 1", value: hf("", "Page &P", "") },
    { label: "Page 1 of ?", value: hf("", "Page &P of &N", "") },
    { label: sheetName, value: hf("", "&A", "") },
    { label: `Confidential, ${new Date().toLocaleDateString()}, Page 1`, value: hf("Confidential", "&D", "Page &P") },
    { label: fileName, value: hf("", "&F", "") },
    { label: `${sheetName}, Page 1`, value: hf("", "&A", "Page &P") },
    { label: `${sheetName}, Confidential, Page 1`, value: hf("&A", "Confidential", "Page &P") },
    { label: `${fileName}, Page 1`, value: hf("", "&F", "Page &P") },
    { label: `Page 1, ${sheetName}`, value: hf("", "Page &P", "&A") },
    { label: `Page 1, ${fileName}`, value: hf("", "Page &P", "&F") },
    { label: `${fileName}, ${sheetName}`, value: hf("", "&F", "&A") },
  ]
}

/** Excel shows codes by name in Custom Header (&[Page]) and keeps them as letters in the file (&P). */
const CODE_NAMES: [string, string][] = [
  ["&P", "&[Page]"],
  ["&N", "&[Pages]"],
  ["&D", "&[Date]"],
  ["&T", "&[Time]"],
  ["&Z", "&[Path]"],
  ["&F", "&[File]"],
  ["&A", "&[Tab]"],
]
export const codesToNames = (s: string) =>
  s.replace(/&&|&[PNDTZFA]/g, (code) => CODE_NAMES.find(([c]) => c === code)?.[1] ?? code)
export const namesToCodes = (s: string) => CODE_NAMES.reduce((out, [code, name]) => out.split(name).join(code), s)

const sameHF = (a: HeaderFooterText, b: HeaderFooterText) =>
  a.left === b.left && a.center === b.center && a.right === b.right

/** Range text for the Sheet tab ("$A$1:$H$60", "$1:$2", "$A:$A"). */
function rangeText(r: Range): string {
  return `$${columnName(r.c0)}$${r.r0 + 1}:$${columnName(r.c1)}$${r.r1 + 1}`
}
const linesText = (span: [number, number] | null, cols: boolean) =>
  span ? (cols ? `$${columnName(span[0])}:$${columnName(span[1])}` : `$${span[0] + 1}:$${span[1] + 1}`) : ""

function parseLines(text: string, cols: boolean): [number, number] | null | undefined {
  const t = text.replace(/\$/g, "").replace(/^.*!/, "").trim()
  if (!t) return null
  const parts = t.split(":")
  const one = (p: string) => {
    if (cols) {
      if (!/^[A-Z]+$/i.test(p)) return -1
      return parseRange(p.toUpperCase() + "1")?.c0 ?? -1
    }
    return /^\d+$/.test(p) ? Number(p) - 1 : -1
  }
  const a = one(parts[0])
  const b = parts.length > 1 ? one(parts[1]) : a
  if (a < 0 || b < 0 || parts.length > 2) return undefined
  return [Math.min(a, b), Math.max(a, b)]
}

function parseArea(text: string): Range[] | null | undefined {
  const t = text.trim()
  if (!t) return null
  const out: Range[] = []
  for (const part of t.split(",")) {
    const r = parseRange(part.replace(/\$/g, "").replace(/^.*!/, "").trim().toUpperCase())
    if (!r) return undefined
    out.push(r)
  }
  return out
}

/**
 * Page Layout › Page Setup (the group's launchers, Print Titles, File › Print's Page Setup link): Excel's four tabs.
 * Changes apply together on OK, as one undo step.
 */
export function PageSetupDialog({
  controller,
  tab: initial = "page",
  fileName,
  onClose,
}: {
  controller: SheetController
  tab?: PageSetupTab
  fileName: string
  onClose: () => void
}) {
  const sheet = controller.sheet
  const [tab, setTab] = useState<PageSetupTab>(initial)
  const [setup, setSetup] = useState<PageSetup>(() => copyPageSetup(sheet.pageSetup))
  const set = (fields: Partial<PageSetup>) => setSetup((s) => ({ ...s, ...fields }))
  const [area, setArea] = useState(() => (setup.printArea ?? []).map(rangeText).join(","))
  const [rows, setRows] = useState(() => linesText(setup.titleRows, false))
  const [cols, setCols] = useState(() => linesText(setup.titleCols, true))
  const [custom, setCustom] = useState<"header" | "footer" | null>(null)

  const accept = () => {
    const printArea = parseArea(area)
    const titleRows = parseLines(rows, false)
    const titleCols = parseLines(cols, true)
    if (printArea === undefined || titleRows === undefined || titleCols === undefined) {
      controller.shell.alert(BAD_REFERENCE)
      setTab("sheet")
      return
    }
    controller.print.change((s) => {
      Object.assign(s, copyPageSetup(setup))
      s.printArea = printArea
      s.titleRows = titleRows
      s.titleCols = titleCols
    })
    onClose()
  }

  return (
    <Dialog title="Page Setup" onCancel={onClose} onAccept={accept} width={500}>
      <div className="fc-tabs" role="tablist">
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            aria-selected={tab === t.id}
            className="fc-tab"
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className="fc-page ps-page" role="tabpanel" aria-label={TABS.find((t) => t.id === tab)!.label}>
        {tab === "page" ? <PageTab controller={controller} setup={setup} set={set} /> : null}
        {tab === "margins" ? <MarginsTab controller={controller} setup={setup} set={set} /> : null}
        {tab === "headerFooter" ? (
          <HeaderFooterTab
            setup={setup}
            set={set}
            sheet={sheet}
            fileName={fileName}
            onCustom={(which) => setCustom(which)}
          />
        ) : null}
        {tab === "sheet" ? (
          <SheetTab
            setup={setup}
            set={set}
            area={area}
            setArea={setArea}
            rows={rows}
            setRows={setRows}
            cols={cols}
            setCols={setCols}
          />
        ) : null}
      </div>
      {custom ? (
        <CustomHeaderFooterDialog
          title={custom === "header" ? "Header" : "Footer"}
          value={setup[custom]}
          onCancel={() => setCustom(null)}
          onAccept={(value) => {
            set({ [custom]: value })
            setCustom(null)
          }}
        />
      ) : null}
    </Dialog>
  )
}

type TabProps = { setup: PageSetup; set: (fields: Partial<PageSetup>) => void }

function Radio({
  name,
  label,
  checked,
  onChange,
}: {
  name: string
  label: ReactNode
  checked: boolean
  onChange: () => void
}) {
  return (
    <label className="fc-check">
      <input type="radio" name={name} checked={checked} onChange={onChange} />
      <span>{label}</span>
    </label>
  )
}

function Check({ label, checked, onChange }: { label: string; checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <label className="fc-check">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

function NumberBox({
  value,
  onChange,
  min,
  max,
  step = 1,
  disabled,
  label,
  width = 64,
}: {
  value: number
  onChange: (v: number) => void
  min: number
  max: number
  step?: number
  disabled?: boolean
  label: string
  width?: number
}) {
  return (
    <input
      type="number"
      className="dialog-field ps-number"
      style={{ width }}
      aria-label={label}
      value={Number(value.toFixed(2))}
      min={min}
      max={max}
      step={step}
      disabled={disabled}
      onChange={(e) => {
        const v = Number(e.target.value)
        if (Number.isFinite(v)) onChange(Math.min(max, Math.max(min, v)))
      }}
    />
  )
}

function PageTab({ controller, setup, set }: TabProps & { controller: SheetController }) {
  const fit = setup.fitTo
  const paper = paperOf(setup, controller.print.defaultPaper)
  return (
    <div className="ps-column">
      <fieldset className="dialog-group">
        <legend>Orientation</legend>
        <div className="ps-row">
          <Radio
            name="ps-orient"
            label="Portrait"
            checked={setup.orientation === "portrait"}
            onChange={() => set({ orientation: "portrait" })}
          />
          <Radio
            name="ps-orient"
            label="Landscape"
            checked={setup.orientation === "landscape"}
            onChange={() => set({ orientation: "landscape" })}
          />
        </div>
      </fieldset>
      <fieldset className="dialog-group">
        <legend>Scaling</legend>
        <div className="ps-row">
          <Radio name="ps-scale" label="Adjust to:" checked={!fit} onChange={() => set({ fitTo: null })} />
          <NumberBox
            label="Adjust to percent"
            value={setup.scale}
            min={10}
            max={400}
            disabled={!!fit}
            onChange={(v) => set({ scale: Math.round(v) })}
          />
          <span>% normal size</span>
        </div>
        <div className="ps-row">
          <Radio
            name="ps-scale"
            label="Fit to:"
            checked={!!fit}
            onChange={() => set({ fitTo: { width: 1, height: 1 } })}
          />
          <NumberBox
            label="Pages wide"
            value={fit?.width ?? 1}
            min={0}
            max={32767}
            disabled={!fit}
            onChange={(v) => set({ fitTo: { width: Math.round(v), height: fit?.height ?? 1 } })}
          />
          <span>page(s) wide by</span>
          <NumberBox
            label="Pages tall"
            value={fit?.height ?? 1}
            min={0}
            max={32767}
            disabled={!fit}
            onChange={(v) => set({ fitTo: { width: fit?.width ?? 1, height: Math.round(v) } })}
          />
          <span>tall</span>
        </div>
      </fieldset>
      <label className="ps-line">
        <span className="dialog-label">
          Paper si<u>z</u>e:
        </span>
        <select className="dialog-field" value={paper} onChange={(e) => set({ paperSize: Number(e.target.value) })}>
          {Object.entries(PAPER_SIZES).map(([id, p]) => (
            <option key={id} value={id}>
              {p.name}
            </option>
          ))}
        </select>
      </label>
      <label className="ps-line">
        <span className="dialog-label">
          Fi<u>r</u>st page number:
        </span>
        <input
          className="dialog-field ps-number"
          style={{ width: 90 }}
          value={setup.firstPageNumber === null ? "Auto" : String(setup.firstPageNumber)}
          onChange={(e) => {
            const t = e.target.value.trim()
            if (!t || /^auto$/i.test(t)) set({ firstPageNumber: null })
            else if (/^-?\d+$/.test(t)) set({ firstPageNumber: Number(t) })
          }}
        />
      </label>
    </div>
  )
}

function MarginsTab({ controller, setup, set }: TabProps & { controller: SheetController }) {
  const unit = marginUnit(controller)
  const m = setup.margins
  const box = (key: keyof typeof m, label: string) => (
    <label className="ps-margin">
      <span className="dialog-label">{label}</span>
      <NumberBox
        label={label.replace(":", "")}
        value={m[key] * unit.perInch}
        min={0}
        max={50}
        step={unit.step}
        onChange={(v) => set({ margins: { ...m, [key]: v / unit.perInch } })}
      />
    </label>
  )
  return (
    <div className="ps-column">
      <div className="ps-margins">
        <div className="ps-margins-row">
          {box("top", "Top:")}
          {box("header", "Header:")}
        </div>
        <div className="ps-margins-row ps-margins-middle">
          {box("left", "Left:")}
          <div className="ps-paper" data-keep-light aria-hidden>
            <div
              className="ps-paper-body"
              style={{ justifyContent: setup.centerHorizontally ? "center" : "flex-start" }}
            />
          </div>
          {box("right", "Right:")}
        </div>
        <div className="ps-margins-row">
          {box("bottom", "Bottom:")}
          {box("footer", "Footer:")}
        </div>
      </div>
      <p className="ps-note">Margins are in {unit.name === "in" ? "inches" : "centimetres"}.</p>
      <fieldset className="dialog-group">
        <legend>Center on page</legend>
        <div className="ps-row">
          <Check
            label="Horizontally"
            checked={setup.centerHorizontally}
            onChange={(v) => set({ centerHorizontally: v })}
          />
          <Check label="Vertically" checked={setup.centerVertically} onChange={(v) => set({ centerVertically: v })} />
        </div>
      </fieldset>
    </div>
  )
}

function HeaderFooterTab({
  setup,
  set,
  sheet,
  fileName,
  onCustom,
}: TabProps & { sheet: Sheet; fileName: string; onCustom: (which: "header" | "footer") => void }) {
  const presets = headerPresets(sheet.name, fileName)
  // The previews show what prints on page 1, as Excel's do.
  const ctx: PrintContext = { fileName, folder: "", now: new Date(), total: 1 }
  const show = (s: string) => headerFooterLines(s, PREVIEW_PAGE, sheet, ctx).join("\n")
  const pick = (which: "header" | "footer") => {
    const value = setup[which]
    const index = presets.findIndex((p) => sameHF(p.value, value))
    return (
      <select
        className="dialog-field"
        aria-label={which === "header" ? "Header" : "Footer"}
        value={index < 0 ? "custom" : String(index)}
        onChange={(e) => {
          const p = presets[Number(e.target.value)]
          if (p) set({ [which]: { ...p.value } })
        }}
      >
        {index < 0 ? <option value="custom">(custom)</option> : null}
        {presets.map((p, i) => (
          <option key={i} value={i}>
            {p.label}
          </option>
        ))}
      </select>
    )
  }
  return (
    <div className="ps-column">
      <HFPreview value={setup.header} show={show} />
      <label className="ps-line ps-stack">
        <span className="dialog-label">
          He<u>a</u>der:
        </span>
        {pick("header")}
      </label>
      <div className="ps-row ps-buttons">
        <button type="button" className="dialog-btn" onClick={() => onCustom("header")}>
          <u>C</u>ustom Header...
        </button>
        <button type="button" className="dialog-btn" onClick={() => onCustom("footer")}>
          C<u>u</u>stom Footer...
        </button>
      </div>
      <label className="ps-line ps-stack">
        <span className="dialog-label">
          <u>F</u>ooter:
        </span>
        {pick("footer")}
      </label>
      <HFPreview value={setup.footer} show={show} />
    </div>
  )
}

const PREVIEW_PAGE: PrintPage = { area: 0, rows: [0, 0], cols: [0, 0], titleRows: null, titleCols: null, number: 1 }

function HFPreview({ value, show }: { value: HeaderFooterText; show: (s: string) => string }) {
  return (
    <div className="ps-hf-preview" data-keep-light aria-label="Preview">
      <span>{show(value.left)}</span>
      <span>{show(value.center)}</span>
      <span>{show(value.right)}</span>
    </div>
  )
}

function SheetTab({
  setup,
  set,
  area,
  setArea,
  rows,
  setRows,
  cols,
  setCols,
}: TabProps & {
  area: string
  setArea: (v: string) => void
  rows: string
  setRows: (v: string) => void
  cols: string
  setCols: (v: string) => void
}) {
  return (
    <div className="ps-column">
      <label className="ps-line">
        <span className="dialog-label">
          Print <u>a</u>rea:
        </span>
        <input className="dialog-field" value={area} spellCheck={false} onChange={(e) => setArea(e.target.value)} />
      </label>
      <fieldset className="dialog-group">
        <legend>Print titles</legend>
        <label className="ps-line">
          <span className="dialog-label">
            <u>R</u>ows to repeat at top:
          </span>
          <input className="dialog-field" value={rows} spellCheck={false} onChange={(e) => setRows(e.target.value)} />
        </label>
        <label className="ps-line">
          <span className="dialog-label">
            <u>C</u>olumns to repeat at left:
          </span>
          <input className="dialog-field" value={cols} spellCheck={false} onChange={(e) => setCols(e.target.value)} />
        </label>
      </fieldset>
      <fieldset className="dialog-group">
        <legend>Print</legend>
        <div className="ps-grid">
          <Check label="Gridlines" checked={setup.gridlines} onChange={(v) => set({ gridlines: v })} />
          <Check label="Black and white" checked={setup.blackAndWhite} onChange={(v) => set({ blackAndWhite: v })} />
          <Check label="Row and column headings" checked={setup.headings} onChange={(v) => set({ headings: v })} />
        </div>
      </fieldset>
      <fieldset className="dialog-group">
        <legend>Page order</legend>
        <Radio
          name="ps-order"
          label="Down, then over"
          checked={setup.order === "downThenOver"}
          onChange={() => set({ order: "downThenOver" })}
        />
        <Radio
          name="ps-order"
          label="Over, then down"
          checked={setup.order === "overThenDown"}
          onChange={() => set({ order: "overThenDown" })}
        />
      </fieldset>
    </div>
  )
}

/** Custom Header / Custom Footer: the three sections, with buttons that insert Excel's codes. */
function CustomHeaderFooterDialog({
  title,
  value,
  onCancel,
  onAccept,
}: {
  title: string
  value: HeaderFooterText
  onCancel: () => void
  onAccept: (value: HeaderFooterText) => void
}) {
  const [text, setText] = useState(() => ({
    left: codesToNames(value.left),
    center: codesToNames(value.center),
    right: codesToNames(value.right),
  }))
  const [focus, setFocus] = useState<keyof HeaderFooterText>("left")
  const insert = (name: string) => setText((t) => ({ ...t, [focus]: t[focus] + name }))
  const BUTTONS: [string, string][] = [
    ["Insert Page Number", "&[Page]"],
    ["Insert Number of Pages", "&[Pages]"],
    ["Insert Date", "&[Date]"],
    ["Insert Time", "&[Time]"],
    ["Insert File Path", "&[Path]&[File]"],
    ["Insert File Name", "&[File]"],
    ["Insert Sheet Name", "&[Tab]"],
  ]
  return (
    <Dialog
      title={title}
      width={560}
      onCancel={onCancel}
      onAccept={() =>
        onAccept({
          left: namesToCodes(text.left),
          center: namesToCodes(text.center),
          right: namesToCodes(text.right),
        })
      }
    >
      <p className="ps-note">
        To insert a page number, date, time, file path, filename, or tab name: position the insertion point in the edit
        box, then choose the appropriate button.
      </p>
      <div className="ps-row ps-codes">
        {BUTTONS.map(([label, name]) => (
          <button key={label} type="button" className="dialog-btn" title={label} onClick={() => insert(name)}>
            {label.replace("Insert ", "")}
          </button>
        ))}
      </div>
      <div className="ps-sections">
        {(["left", "center", "right"] as const).map((k) => (
          <label key={k} className="ps-section">
            <span className="dialog-label">
              {k === "left" ? "Left section:" : k === "center" ? "Center section:" : "Right section:"}
            </span>
            <textarea
              className="dialog-field"
              rows={4}
              value={text[k]}
              onFocus={() => setFocus(k)}
              onChange={(e) => setText((t) => ({ ...t, [k]: e.target.value }))}
            />
          </label>
        ))}
      </div>
    </Dialog>
  )
}
