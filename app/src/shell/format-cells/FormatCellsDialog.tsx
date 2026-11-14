import {
  formatValue,
  formatWithCode,
  type Border,
  type BorderStyle,
  type CellStyle,
  type HAlign,
  type Op,
  type StylePatch,
  type VAlign,
} from "@cells/core"
import { useMemo, useState, type ReactNode } from "react"
import { FONT_NAMES } from "@/ribbon/home-commands"
import { ColorPanel } from "@/ribbon/colors"
import { cn } from "@/lib/cn"
import type { SheetController } from "@/sheet/controller"
import { DEFAULT_FONT } from "@/sheet/format"
import { Popover } from "@/ui/popover"
import { Dialog } from "../Dialog"
import { borderPlanOps, readBorders, type BorderPlan, type Edge } from "./borders"
import {
  CATEGORIES,
  CATEGORY_HELP,
  CUSTOM_TYPES,
  DATE_TYPES,
  FRACTION_TYPES,
  SPECIAL_TYPES,
  TIME_TYPES,
  buildCode,
  readCode,
  type Category,
  type Negative,
} from "./number-formats"

export type FormatTab = "number" | "alignment" | "font" | "border" | "fill" | "protection"

const TABS: { id: FormatTab; label: string }[] = [
  { id: "number", label: "Number" },
  { id: "alignment", label: "Alignment" },
  { id: "font", label: "Font" },
  { id: "border", label: "Border" },
  { id: "fill", label: "Fill" },
  { id: "protection", label: "Protection" },
]

/**
 * Format Cells (Ctrl+1), Excel's six tabs. It opens on the active cell's format and applies only what was changed,
 * as one undo step, so formats the dialog merely displayed stay as they were on every other cell.
 */
export function FormatCellsDialog({
  controller,
  tab: initialTab = "number",
  onClose,
}: {
  controller: SheetController
  tab?: FormatTab
  onClose: () => void
}) {
  const [tab, setTab] = useState<FormatTab>(initialTab)
  const start = controller.format.active
  const [patch, setPatch] = useState<StylePatch>({})
  const [borders, setBorders] = useState<BorderPlan>({})
  const [merge, setMerge] = useState<boolean | null>(null)
  const style: CellStyle = { ...start, ...(patch as CellStyle) }
  const set = (fields: StylePatch) => setPatch((p) => ({ ...p, ...fields }))

  const range = controller.selection.ranges[controller.selection.ranges.length - 1]
  const initialBorders = useMemo(() => readBorders(controller.sheet, controller.wb.styles, range), [controller, range])

  const accept = () => {
    const sheet = controller.sheetIndex
    const ops: Op[] = []
    // Drop fields set back to what they were (no direct formatting for a value only displayed).
    const changed = Object.fromEntries(
      Object.entries(patch).filter(([k, v]) => (start as Record<string, unknown>)[k] !== (v ?? undefined)),
    ) as StylePatch
    if (Object.keys(changed).length)
      ops.push({ kind: "format", sheet, ranges: [...controller.selection.ranges], patch: changed })
    ops.push(...borderPlanOps(sheet, range, borders))
    if (merge === true && !controller.format.isMerged()) ops.push({ kind: "merge", sheet, ranges: [range] })
    if (merge === false && controller.format.isMerged())
      ops.push({ kind: "unmerge", sheet, ranges: [...controller.selection.ranges] })
    if (ops.length) {
      const sel = controller.selection
      controller.transact(ops, { label: "Format Cells", before: sel, after: sel })
    }
    onClose()
  }

  return (
    <Dialog title="Format Cells" onCancel={onClose} onAccept={accept} width={560}>
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
      <div className="fc-page" role="tabpanel" aria-label={TABS.find((t) => t.id === tab)!.label}>
        {tab === "number" ? <NumberTab controller={controller} style={style} set={set} /> : null}
        {tab === "alignment" ? (
          <AlignmentTab style={style} set={set} merged={merge ?? controller.format.isMerged()} setMerge={setMerge} />
        ) : null}
        {tab === "font" ? <FontTab style={style} set={set} /> : null}
        {tab === "border" ? (
          <BorderTab
            plan={{ ...initialBorders, ...borders }}
            multiRow={range.r1 > range.r0}
            multiCol={range.c1 > range.c0}
            onChange={(edges) => setBorders((b) => ({ ...b, ...edges }))}
          />
        ) : null}
        {tab === "fill" ? <FillTab style={style} set={set} /> : null}
        {tab === "protection" ? <ProtectionTab style={style} set={set} /> : null}
      </div>
    </Dialog>
  )
}

type TabProps = { style: CellStyle; set: (fields: StylePatch) => void }

function Field({ label, children, className }: { label: string; children: ReactNode; className?: string }) {
  return (
    <label className={cn("fc-field", className)}>
      <span className="dialog-label">{label}</span>
      {children}
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

function List<T extends string>({
  label,
  items,
  value,
  onChange,
  render,
  className,
}: {
  label: string
  items: readonly T[]
  value: T | undefined
  onChange: (v: T) => void
  render?: (v: T) => ReactNode
  className?: string
}) {
  return (
    <div className={cn("dialog-list fc-list", className)} role="listbox" aria-label={label}>
      {items.map((item) => (
        <button
          key={item}
          type="button"
          role="option"
          aria-selected={item === value}
          className={cn("dialog-list-item", item === value && "is-selected")}
          onClick={() => onChange(item)}
        >
          {render ? render(item) : item}
        </button>
      ))}
    </div>
  )
}

// ── Number ───────────────────────────────────────────────────────────────────────────────────────────────────

function NumberTab({ controller, style, set }: TabProps & { controller: SheetController }) {
  // What the dialog opened with (style changes while it is open come from the dialog itself).
  const [opened] = useState(() => readCode(style.numberFormat))
  const [category, setCategory] = useState<Category>(opened.category)
  const [options, setOptions] = useState(opened.options)
  const [type, setType] = useState<string | undefined>(opened.type ?? style.numberFormat)
  const { row, col } = controller.selection.active
  const value = controller.sheet.cells.value(row, col)

  const apply = (c: Category, o = options, t = type) => {
    const code = buildCode(c, o, t)
    set({ numberFormat: code === "General" ? null : code })
  }
  const pick = (c: Category) => {
    setCategory(c)
    const t = c === "Custom" ? (style.numberFormat ?? "General") : undefined
    setType(t)
    apply(c, options, t)
  }
  const option = (o: Partial<typeof options>) => {
    const next = { ...options, ...o }
    setOptions(next)
    apply(category, next, type)
  }
  const chooseType = (t: string) => {
    setType(t)
    apply(category, options, t)
  }

  const code = buildCode(category, options, type)
  const sample =
    value === undefined
      ? ""
      : code === "General"
        ? (formatValue(value) ?? "")
        : typeof value === "object"
          ? (formatValue(value) ?? "")
          : formatWithCode(value, code).text
  const negatives = (["-1,234.10", "1,234.10", "(1,234.10)", "(1,234.10)"] as const).map((text, i) => ({
    text: category === "Currency" && options.symbol ? text.replace("1", `${options.symbol}1`) : text,
    red: i === 1 || i === 3,
  }))

  return (
    <div className="fc-number">
      <div className="fc-column">
        <span className="dialog-label">Category:</span>
        <List label="Category" items={CATEGORIES} value={category} onChange={pick} className="fc-categories" />
      </div>
      <div className="fc-column fc-grow">
        <fieldset className="dialog-group fc-sample">
          <legend>Sample</legend>
          <span>{sample || " "}</span>
        </fieldset>
        {["Number", "Currency", "Accounting", "Percentage", "Scientific"].includes(category) ? (
          <Field label="Decimal places:" className="is-inline">
            <input
              type="number"
              className="dialog-field fc-spin"
              min={0}
              max={30}
              value={options.decimals}
              onChange={(e) => option({ decimals: Math.max(0, Math.min(30, Number(e.target.value) || 0)) })}
            />
          </Field>
        ) : null}
        {category === "Number" ? (
          <Check
            label="Use 1000 Separator (,)"
            checked={options.separator}
            onChange={(v) => option({ separator: v })}
          />
        ) : null}
        {category === "Currency" || category === "Accounting" ? (
          <Field label="Symbol:" className="is-inline">
            <select
              className="dialog-field"
              value={options.symbol}
              onChange={(e) => option({ symbol: e.target.value })}
            >
              <option value="">None</option>
              <option value="$">$</option>
            </select>
          </Field>
        ) : null}
        {category === "Number" || category === "Currency" ? (
          <>
            <span className="dialog-label">Negative numbers:</span>
            <div className="dialog-list fc-list fc-negatives" role="listbox" aria-label="Negative numbers">
              {negatives.map((n, i) => (
                <button
                  key={i}
                  type="button"
                  role="option"
                  aria-selected={options.negative === i}
                  className={cn("dialog-list-item", options.negative === i && "is-selected")}
                  style={n.red ? { color: "#e00000" } : undefined}
                  onClick={() => option({ negative: i as Negative })}
                >
                  {n.text}
                </button>
              ))}
            </div>
          </>
        ) : null}
        {category === "Date" || category === "Time" ? (
          <>
            <span className="dialog-label">Type:</span>
            <List
              label="Type"
              items={category === "Date" ? DATE_TYPES : TIME_TYPES}
              value={type}
              onChange={chooseType}
              render={(t) => formatWithCode(category === "Date" ? 46036.5625 : 46036.5625, t).text}
            />
          </>
        ) : null}
        {category === "Fraction" || category === "Special" ? (
          <>
            <span className="dialog-label">Type:</span>
            <List
              label="Type"
              items={(category === "Fraction" ? FRACTION_TYPES : SPECIAL_TYPES).map((t) => t.code)}
              value={type}
              onChange={chooseType}
              render={(t) => [...FRACTION_TYPES, ...SPECIAL_TYPES].find((x) => x.code === t)!.label}
            />
          </>
        ) : null}
        {category === "Custom" ? (
          <>
            <Field label="Type:">
              <input
                className="dialog-field"
                value={type ?? ""}
                spellCheck={false}
                onChange={(e) => chooseType(e.target.value)}
              />
            </Field>
            <List label="Format codes" items={CUSTOM_TYPES} value={type} onChange={chooseType} />
          </>
        ) : null}
        <p className="fc-help">{CATEGORY_HELP[category]}</p>
      </div>
    </div>
  )
}

// ── Alignment ────────────────────────────────────────────────────────────────────────────────────────────────

const H_ALIGNS: { value: HAlign; label: string }[] = [
  { value: "general", label: "General" },
  { value: "left", label: "Left (Indent)" },
  { value: "center", label: "Center" },
  { value: "right", label: "Right (Indent)" },
  { value: "fill", label: "Fill" },
  { value: "justify", label: "Justify" },
  { value: "centerContinuous", label: "Center Across Selection" },
  { value: "distributed", label: "Distributed (Indent)" },
]

const V_ALIGNS: { value: VAlign; label: string }[] = [
  { value: "top", label: "Top" },
  { value: "center", label: "Center" },
  { value: "bottom", label: "Bottom" },
  { value: "justify", label: "Justify" },
  { value: "distributed", label: "Distributed" },
]

/** Excel's rotation encoding to degrees (−90…90) and back. */
const toDegrees = (r: number | undefined) => (!r || r === 255 ? 0 : r <= 90 ? r : 90 - r)
const fromDegrees = (d: number) => (d >= 0 ? d : 90 - d)

function AlignmentTab({
  style,
  set,
  merged,
  setMerge,
}: TabProps & { merged: boolean; setMerge: (v: boolean) => void }) {
  const h = style.hAlign ?? "general"
  const indentable = h === "left" || h === "right" || h === "distributed"
  const degrees = toDegrees(style.rotation)
  return (
    <div className="fc-two">
      <div className="fc-column fc-grow">
        <fieldset className="dialog-group">
          <legend>Text alignment</legend>
          <div className="fc-row">
            <Field label="Horizontal:">
              <select
                className="dialog-field"
                value={h}
                onChange={(e) => {
                  const v = e.target.value as HAlign
                  set({
                    hAlign: v === "general" ? null : v,
                    indent: v === "left" || v === "right" || v === "distributed" ? undefined : null,
                  })
                }}
              >
                {H_ALIGNS.map((a) => (
                  <option key={a.value} value={a.value}>
                    {a.label}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Indent:">
              <input
                type="number"
                className="dialog-field fc-spin"
                min={0}
                max={250}
                disabled={!indentable}
                value={style.indent ?? 0}
                onChange={(e) => set({ indent: Number(e.target.value) || null })}
              />
            </Field>
          </div>
          <Field label="Vertical:">
            <select
              className="dialog-field"
              value={style.vAlign ?? "bottom"}
              onChange={(e) => {
                const v = e.target.value as VAlign
                set({ vAlign: v === "bottom" ? null : v })
              }}
            >
              {V_ALIGNS.map((a) => (
                <option key={a.value} value={a.value}>
                  {a.label}
                </option>
              ))}
            </select>
          </Field>
        </fieldset>
        <fieldset className="dialog-group">
          <legend>Text control</legend>
          <Check label="Wrap text" checked={!!style.wrap} onChange={(v) => set({ wrap: v || null })} />
          <Check label="Shrink to fit" checked={!!style.shrink} onChange={(v) => set({ shrink: v || null })} />
          <Check label="Merge cells" checked={merged} onChange={setMerge} />
        </fieldset>
      </div>
      <fieldset className="dialog-group fc-orientation">
        <legend>Orientation</legend>
        <button
          type="button"
          className={cn("fc-stacked", style.rotation === 255 && "is-on")}
          aria-pressed={style.rotation === 255}
          onClick={() => set({ rotation: style.rotation === 255 ? null : 255 })}
        >
          T<br />e<br />x<br />t
        </button>
        <div className="fc-dial" aria-hidden>
          <span style={{ rotate: `${-degrees}deg` }}>Text ——</span>
        </div>
        <Field label="Degrees" className="is-inline">
          <input
            type="number"
            className="dialog-field fc-spin"
            min={-90}
            max={90}
            value={degrees}
            onChange={(e) => {
              const d = Math.max(-90, Math.min(90, Math.round(Number(e.target.value) || 0)))
              set({ rotation: d === 0 ? null : fromDegrees(d) })
            }}
          />
        </Field>
      </fieldset>
    </div>
  )
}

// ── Font ─────────────────────────────────────────────────────────────────────────────────────────────────────

const FONT_STYLES = ["Regular", "Italic", "Bold", "Bold Italic"] as const
const SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"]

function FontTab({ style, set }: TabProps) {
  const font = style.fontName ?? DEFAULT_FONT
  const fontStyle = FONT_STYLES[(style.bold ? 2 : 0) + (style.italic ? 1 : 0)]
  const size = String(style.fontSize ?? 11)
  const [colorOpen, setColorOpen] = useState(false)
  return (
    <div className="fc-font">
      <div className="fc-row fc-font-lists">
        <div className="fc-column fc-grow">
          <Field label="Font:">
            <input className="dialog-field" value={font} onChange={(e) => set({ fontName: e.target.value || null })} />
          </Field>
          <List
            label="Font"
            items={FONT_NAMES}
            value={font}
            onChange={(f) => set({ fontName: f === DEFAULT_FONT ? null : f })}
            render={(f) => <span style={{ fontFamily: `"${f}"` }}>{f}</span>}
          />
        </div>
        <div className="fc-column">
          <span className="dialog-label">Font style:</span>
          <List
            label="Font style"
            items={FONT_STYLES}
            value={fontStyle}
            onChange={(s) => set({ bold: s.startsWith("Bold") || null, italic: s.endsWith("Italic") || null })}
          />
        </div>
        <div className="fc-column fc-size">
          <span className="dialog-label">Size:</span>
          <List
            label="Size"
            items={SIZES}
            value={size}
            onChange={(s) => set({ fontSize: s === "11" ? null : Number(s) })}
          />
        </div>
      </div>
      <div className="fc-row">
        <Field label="Underline:">
          <select
            className="dialog-field"
            value={style.underline ?? "none"}
            onChange={(e) =>
              set({ underline: e.target.value === "none" ? null : (e.target.value as CellStyle["underline"]) })
            }
          >
            <option value="none">None</option>
            <option value="single">Single</option>
            <option value="double">Double</option>
            <option value="singleAccounting">Single Accounting</option>
            <option value="doubleAccounting">Double Accounting</option>
          </select>
        </Field>
        <Field label="Color:">
          <Popover
            open={colorOpen}
            onOpenChange={setColorOpen}
            trigger={
              <button type="button" className="dialog-field fc-color" aria-label="Color">
                <span style={{ background: style.fontColor ?? "#000000" }} />
                {style.fontColor ? "" : "Automatic"}
              </button>
            }
          >
            <ColorPanel
              noneLabel="Automatic"
              current={style.fontColor}
              onPick={(hex) => {
                set({ fontColor: hex })
                setColorOpen(false)
              }}
            />
          </Popover>
        </Field>
      </div>
      <div className="fc-row">
        <fieldset className="dialog-group">
          <legend>Effects</legend>
          <Check label="Strikethrough" checked={!!style.strike} onChange={(v) => set({ strike: v || null })} />
          <Check
            label="Superscript"
            checked={style.vertAlign === "superscript"}
            onChange={(v) => set({ vertAlign: v ? "superscript" : null })}
          />
          <Check
            label="Subscript"
            checked={style.vertAlign === "subscript"}
            onChange={(v) => set({ vertAlign: v ? "subscript" : null })}
          />
        </fieldset>
        <fieldset className="dialog-group fc-grow">
          <legend>Preview</legend>
          <div
            className="fc-preview"
            data-keep-light
            style={{
              fontFamily: `"${font}", "Aptos Narrow", Carlito, sans-serif`,
              fontSize: `${Math.min(28, Number(size))}pt`,
              fontWeight: style.bold ? 700 : 400,
              fontStyle: style.italic ? "italic" : "normal",
              textDecoration:
                [style.underline ? "underline" : "", style.strike ? "line-through" : ""].join(" ").trim() || "none",
              textDecorationStyle: style.underline?.startsWith("double") ? "double" : undefined,
              color: style.fontColor ?? "#000000",
            }}
          >
            AaBbCcYyZz
          </div>
        </fieldset>
      </div>
    </div>
  )
}

// ── Border ───────────────────────────────────────────────────────────────────────────────────────────────────

const LINE_STYLES: BorderStyle[] = [
  "hair",
  "dotted",
  "dashDotDot",
  "dashDot",
  "dashed",
  "thin",
  "mediumDashDotDot",
  "slantDashDot",
  "mediumDashDot",
  "mediumDashed",
  "medium",
  "thick",
  "double",
]

function BorderTab({
  plan,
  multiRow,
  multiCol,
  onChange,
}: {
  plan: BorderPlan
  multiRow: boolean
  multiCol: boolean
  onChange: (edges: BorderPlan) => void
}) {
  const [pen, setPen] = useState<Border>({ style: "thin" })
  const [colorOpen, setColorOpen] = useState(false)
  const toggle = (edge: Edge) => {
    const now = plan[edge]
    const on = now && now.style === pen.style && (now.color ?? "") === (pen.color ?? "")
    onChange({ [edge]: on ? null : { ...pen } })
  }
  const line = (edge: Edge, d: string) => {
    const b = plan[edge]
    if (!b) return null
    const width = b.style === "thick" ? 3 : b.style.startsWith("medium") || b.style === "slantDashDot" ? 2 : 1
    const dash = /dash|dot|hair/i.test(b.style)
      ? b.style === "dotted" || b.style === "hair"
        ? "1 2"
        : "4 2"
      : undefined
    return (
      <path
        key={edge}
        d={d}
        stroke={b.color ?? "currentColor"}
        strokeWidth={b.style === "double" ? 1 : width}
        strokeDasharray={dash}
        fill="none"
      />
    )
  }
  const edgeButton = (edge: Edge, label: string, disabled = false) => (
    <button
      type="button"
      className="fc-edge"
      aria-label={label}
      title={label}
      aria-pressed={!!plan[edge]}
      disabled={disabled}
      onClick={() => toggle(edge)}
    >
      <svg viewBox="0 0 16 16" aria-hidden>
        <rect x="2.5" y="2.5" width="11" height="11" stroke="var(--icon-muted)" strokeDasharray="1 1.5" fill="none" />
        <path
          d={
            {
              top: "M2.5 2.5h11",
              bottom: "M2.5 13.5h11",
              left: "M2.5 2.5v11",
              right: "M13.5 2.5v11",
              insideH: "M2.5 8h11",
              insideV: "M8 2.5v11",
              diagUp: "M2.5 13.5 13.5 2.5",
              diagDown: "M2.5 2.5l11 11",
            }[edge]
          }
          stroke="var(--icon-base)"
          strokeWidth={1.4}
        />
      </svg>
    </button>
  )
  return (
    <div className="fc-two">
      <fieldset className="dialog-group fc-line">
        <legend>Line</legend>
        <span className="dialog-label">Style:</span>
        <div className="fc-line-styles" role="listbox" aria-label="Line style">
          <button
            type="button"
            role="option"
            aria-selected={false}
            className="line-style-row"
            onClick={() => onChange({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null })}
          >
            None
          </button>
          {LINE_STYLES.map((s) => (
            <button
              key={s}
              type="button"
              role="option"
              aria-label={s}
              aria-selected={pen.style === s}
              className="line-style-row"
              onClick={() => setPen({ ...pen, style: s })}
            >
              <span className={`line-sample is-${s}`} />
            </button>
          ))}
        </div>
        <span className="dialog-label">Color:</span>
        <Popover
          open={colorOpen}
          onOpenChange={setColorOpen}
          trigger={
            <button type="button" className="dialog-field fc-color" aria-label="Line color">
              <span style={{ background: pen.color ?? "#000000" }} />
              {pen.color ? "" : "Automatic"}
            </button>
          }
        >
          <ColorPanel
            noneLabel="Automatic"
            current={pen.color}
            onPick={(hex) => {
              setPen({ ...pen, color: hex ?? undefined })
              setColorOpen(false)
            }}
          />
        </Popover>
      </fieldset>
      <div className="fc-column fc-grow">
        <span className="dialog-label">Presets</span>
        <div className="fc-row">
          <button
            type="button"
            className="dialog-btn"
            onClick={() => onChange({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null })}
          >
            None
          </button>
          <button
            type="button"
            className="dialog-btn"
            onClick={() => onChange({ top: { ...pen }, bottom: { ...pen }, left: { ...pen }, right: { ...pen } })}
          >
            Outline
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={!multiRow && !multiCol}
            onClick={() =>
              onChange({ insideH: multiRow ? { ...pen } : undefined, insideV: multiCol ? { ...pen } : undefined })
            }
          >
            Inside
          </button>
        </div>
        <span className="dialog-label">Border</span>
        <div className="fc-border-box">
          <div className="fc-edges-v">
            {edgeButton("top", "Top border")}
            {edgeButton("insideH", "Inside horizontal border", !multiRow)}
            {edgeButton("bottom", "Bottom border")}
          </div>
          <svg
            className="fc-border-preview"
            viewBox="0 0 120 90"
            data-keep-light
            role="img"
            aria-label="Border preview"
          >
            <rect x="0" y="0" width="120" height="90" fill="#ffffff" />
            <text x="60" y="49" textAnchor="middle" fontSize="11" fill="#555">
              Text
            </text>
            {line("top", "M10 10h100")}
            {line("bottom", "M10 80h100")}
            {line("left", "M10 10v70")}
            {line("right", "M110 10v70")}
            {multiRow ? line("insideH", "M10 45h100") : null}
            {multiCol ? line("insideV", "M60 10v70") : null}
            {line("diagUp", "M10 80 110 10")}
            {line("diagDown", "M10 10l100 70")}
          </svg>
          <div className="fc-edges-h">
            {edgeButton("diagUp", "Diagonal up border")}
            {edgeButton("left", "Left border")}
            {edgeButton("insideV", "Inside vertical border", !multiCol)}
            {edgeButton("right", "Right border")}
            {edgeButton("diagDown", "Diagonal down border")}
          </div>
        </div>
        <p className="fc-help">
          The selected border style can be applied by clicking the presets, preview diagram or the buttons above.
        </p>
      </div>
    </div>
  )
}

// ── Fill ─────────────────────────────────────────────────────────────────────────────────────────────────────

function FillTab({ style, set }: TabProps) {
  return (
    <div className="fc-two">
      <div className="fc-column">
        <span className="dialog-label">Background Color:</span>
        <ColorPanel noneLabel="No Color" current={style.fill} onPick={(hex) => set({ fill: hex })} />
      </div>
      <fieldset className="dialog-group fc-grow">
        <legend>Sample</legend>
        <div className="fc-fill-sample" data-keep-light style={{ background: style.fill ?? "#ffffff" }} />
      </fieldset>
    </div>
  )
}

// ── Protection ───────────────────────────────────────────────────────────────────────────────────────────────

function ProtectionTab({ style, set }: TabProps) {
  return (
    <div className="fc-column">
      <Check label="Locked" checked={style.locked !== false} onChange={(v) => set({ locked: v ? null : false })} />
      <Check label="Hidden" checked={!!style.hiddenFormula} onChange={(v) => set({ hiddenFormula: v || null })} />
      <p className="fc-help">
        Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect group,
        Protect Sheet button).
      </p>
    </div>
  )
}
