import {
  NUMBER_FORMATS,
  chartHasAxes,
  chartParts,
  partLabel,
  type ChartSpec,
  type LegendPosition,
  type PartFormat,
} from "@cells/chart"
import { X } from "lucide-react"
import { useEffect, useState, type ReactNode } from "react"
import { ColorPanel } from "@/ribbon/colors"
import { chartPartCan } from "@/sheet/charts"
import type { SheetController } from "@/sheet/controller"
import { Popover } from "@/ui/popover"

/**
 * Excel's Format pane for a chart: the part picked (Chart Area, Plot Area, the title, an axis, a series, the
 * legend…) with its fill, outline and text, and the part's own options. Every change shows at once and is one undo
 * step; a part the pane can't change says so rather than offering what wouldn't show.
 */
export function ChartFormatPane({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  const charts = controller.charts
  const spec = charts.spec()
  const close = () => {
    charts.closeFormatPane()
    onDone()
  }
  if (!spec) return null
  const part = charts.part
  const own = spec.format[part] ?? {}
  const can = chartPartCan(part)
  const label = partLabel(spec, part)
  const format = (change: Partial<PartFormat>, what: string) => charts.formatPart(change, what)
  const set = (change: (s: ChartSpec) => ChartSpec, what = `Format ${label}`) => charts.change(change, what)
  return (
    <aside
      className="comments-pane chart-format-pane"
      aria-label={`Format ${label}`}
      onKeyDown={(e) => {
        if (e.key === "Escape") {
          e.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Format {label}</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <div className="cfp-body">
        <select
          className="dialog-field cfp-parts"
          aria-label="Chart Elements"
          value={part}
          onChange={(e) => charts.setPart(e.target.value)}
        >
          {chartParts(spec).map((p) => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
        </select>
        {can.fill ? (
          <Section title="Fill">
            <Choice
              name="fill"
              value={own.fill === "none" ? "none" : own.fill ? "solid" : "auto"}
              options={[
                ["auto", "Automatic"],
                ["none", "No fill"],
                ["solid", "Solid fill"],
              ]}
              onChange={(v) =>
                format({ fill: v === "none" ? "none" : v === "solid" ? (own.fill ?? "#4472c4") : undefined }, "Fill")
              }
            />
            {own.fill && own.fill !== "none" ? (
              <Row label="Color">
                <ColorButton
                  label="Fill color"
                  color={own.fill}
                  theme={controller.wb.theme}
                  onPick={(c) => format({ fill: c ?? undefined }, "Fill")}
                />
              </Row>
            ) : null}
          </Section>
        ) : null}
        {can.line ? (
          <Section
            title={part.startsWith("series:") || part === "gridlines" || part.startsWith("axis") ? "Line" : "Border"}
          >
            <Choice
              name="line"
              value={own.line === "none" ? "none" : own.line ? "solid" : "auto"}
              options={[
                ["auto", "Automatic"],
                ["none", "No line"],
                ["solid", "Solid line"],
              ]}
              onChange={(v) =>
                format({ line: v === "none" ? "none" : v === "solid" ? (own.line ?? "#595959") : undefined }, "Line")
              }
            />
            {own.line && own.line !== "none" ? (
              <Row label="Color">
                <ColorButton
                  label="Line color"
                  color={own.line}
                  theme={controller.wb.theme}
                  onPick={(c) => format({ line: c ?? undefined }, "Line")}
                />
              </Row>
            ) : null}
            {own.line !== "none" ? (
              <Row label="Width">
                <NumberInput
                  label="Line width"
                  value={own.lineWidth}
                  unit="pt"
                  min={0.25}
                  max={20}
                  step={0.25}
                  onCommit={(lineWidth) => format({ lineWidth }, "Line Width")}
                />
              </Row>
            ) : null}
          </Section>
        ) : null}
        {can.text ? (
          <Section title="Font">
            <Row label="Color">
              <ColorButton
                label="Text color"
                color={own.color}
                automatic
                theme={controller.wb.theme}
                onPick={(c) => format({ color: c ?? undefined }, "Text Fill")}
              />
            </Row>
            <Row label="Size">
              <NumberInput
                label="Font size"
                value={own.fontSize}
                unit="pt"
                min={6}
                max={72}
                step={1}
                onCommit={(fontSize) => format({ fontSize }, "Font Size")}
              />
            </Row>
            <label className="dialog-radio">
              <input
                type="checkbox"
                checked={!!own.bold}
                onChange={(e) => format({ bold: e.target.checked ? true : undefined }, "Bold")}
              />
              <span>Bold</span>
            </label>
          </Section>
        ) : null}
        <PartOptions spec={spec} part={part} set={set} />
      </div>
    </aside>
  )
}

/** The part's own options (Excel's "… Options" page). */
function PartOptions({
  spec,
  part,
  set,
}: {
  spec: ChartSpec
  part: string
  set: (change: (s: ChartSpec) => ChartSpec, what?: string) => void
}) {
  if (part === "title" || part.startsWith("axisTitle:")) {
    const axis = part === "title" ? null : (part.slice("axisTitle:".length) as "category" | "value")
    const text = axis ? spec.axisTitles[axis].text : spec.title.text
    return (
      <Section title="Title Options">
        <Row label="Text">
          <TextInput
            label="Title text"
            value={text}
            onCommit={(t) =>
              set((s) =>
                axis
                  ? { ...s, axisTitles: { ...s.axisTitles, [axis]: { ...s.axisTitles[axis], text: t } } }
                  : { ...s, title: { ...s.title, text: t } },
              )
            }
          />
        </Row>
      </Section>
    )
  }
  if (part === "axis:value") {
    const axis = spec.valueAxis
    const setAxis = (change: Partial<ChartSpec["valueAxis"]>) =>
      set((s) => {
        const next = { ...s.valueAxis, ...change }
        for (const key of Object.keys(next) as (keyof typeof next)[]) if (next[key] === undefined) delete next[key]
        return { ...s, valueAxis: next }
      }, "Format Axis")
    return (
      <Section title="Axis Options">
        <Row label="Minimum">
          <NumberInput label="Minimum" value={axis.min} placeholder="Auto" onCommit={(min) => setAxis({ min })} />
        </Row>
        <Row label="Maximum">
          <NumberInput label="Maximum" value={axis.max} placeholder="Auto" onCommit={(max) => setAxis({ max })} />
        </Row>
        <Row label="Major unit">
          <NumberInput
            label="Major unit"
            value={axis.major}
            placeholder="Auto"
            min={0}
            onCommit={(major) => setAxis({ major: major && major > 0 ? major : undefined })}
          />
        </Row>
        <label className="dialog-radio">
          <input
            type="checkbox"
            checked={!!axis.log}
            onChange={(e) => setAxis({ log: e.target.checked || undefined })}
          />
          <span>Logarithmic scale</span>
        </label>
        <label className="dialog-radio">
          <input
            type="checkbox"
            checked={!!axis.reverse}
            onChange={(e) => setAxis({ reverse: e.target.checked || undefined })}
          />
          <span>Values in reverse order</span>
        </label>
        <Row label="Number">
          <select
            className="dialog-field"
            aria-label="Number format"
            value={axis.numberFormat ?? "General"}
            onChange={(e) => setAxis({ numberFormat: e.target.value === "General" ? undefined : e.target.value })}
          >
            {NUMBER_FORMATS.map((f) => (
              <option key={f.id} value={f.id}>
                {f.label}
              </option>
            ))}
          </select>
        </Row>
      </Section>
    )
  }
  if (part === "axis:category") {
    return (
      <Section title="Axis Options">
        <label className="dialog-radio">
          <input
            type="checkbox"
            checked={spec.categoryReverse}
            onChange={(e) => set((s) => ({ ...s, categoryReverse: e.target.checked }), "Format Axis")}
          />
          <span>Categories in reverse order</span>
        </label>
      </Section>
    )
  }
  if (part === "legend") {
    return (
      <Section title="Legend Options">
        <Choice
          name="legend"
          value={spec.legend}
          options={[
            ["top", "Top"],
            ["bottom", "Bottom"],
            ["left", "Left"],
            ["right", "Right"],
          ]}
          onChange={(legend) => set((s) => ({ ...s, legend: legend as LegendPosition }), "Legend Position")}
        />
      </Section>
    )
  }
  if (part === "dataLabels") {
    const content = spec.labelContent
    const toggle = (key: keyof typeof content) => (on: boolean) =>
      set((s) => ({ ...s, labelContent: { ...s.labelContent, [key]: on } }), "Format Data Labels")
    return (
      <Section title="Label Options">
        <span className="cfp-note">Label contains</span>
        <Check label="Value" checked={content.value} onChange={toggle("value")} />
        <Check label="Category name" checked={content.category} onChange={toggle("category")} />
        <Check label="Series name" checked={content.series} onChange={toggle("series")} />
        {spec.type === "pie" || spec.type === "doughnut" ? (
          <Check label="Percentage" checked={content.percent} onChange={toggle("percent")} />
        ) : null}
      </Section>
    )
  }
  if (part.startsWith("series:") || part.startsWith("point:")) {
    if (spec.type === "column" || spec.type === "bar") {
      return (
        <Section title="Series Options">
          <Row label="Series Overlap">
            <NumberInput
              label="Series Overlap"
              value={spec.overlap}
              unit="%"
              min={-100}
              max={100}
              step={1}
              onCommit={(v) => v !== undefined && set((s) => ({ ...s, overlap: Math.round(v) }), "Series Overlap")}
            />
          </Row>
          <Row label="Gap Width">
            <NumberInput
              label="Gap Width"
              value={spec.gapWidth}
              unit="%"
              min={0}
              max={500}
              step={1}
              onCommit={(v) => v !== undefined && set((s) => ({ ...s, gapWidth: Math.round(v) }), "Gap Width")}
            />
          </Row>
        </Section>
      )
    }
    if (spec.type === "pie" || spec.type === "doughnut") {
      return (
        <Section title="Series Options">
          <Row label="Angle of first slice">
            <NumberInput
              label="Angle of first slice"
              value={spec.firstSliceAngle}
              unit="°"
              min={0}
              max={360}
              step={1}
              onCommit={(v) => v !== undefined && set((s) => ({ ...s, firstSliceAngle: Math.round(v) % 360 }), "Angle")}
            />
          </Row>
          <Row label="Explosion">
            <NumberInput
              label="Pie Explosion"
              value={spec.explosion}
              unit="%"
              min={0}
              max={400}
              step={1}
              onCommit={(v) => v !== undefined && set((s) => ({ ...s, explosion: Math.round(v) }), "Explosion")}
            />
          </Row>
          {spec.type === "doughnut" ? (
            <Row label="Hole Size">
              <NumberInput
                label="Doughnut Hole Size"
                value={spec.holeSize}
                unit="%"
                min={10}
                max={90}
                step={1}
                onCommit={(v) => v !== undefined && set((s) => ({ ...s, holeSize: Math.round(v) }), "Hole Size")}
              />
            </Row>
          ) : null}
        </Section>
      )
    }
    if (spec.type === "line" || spec.type === "scatter") {
      return (
        <Section title="Series Options">
          <Check
            label="Smoothed line"
            checked={spec.smooth}
            onChange={(smooth) => set((s) => ({ ...s, smooth }), "Smoothed Line")}
          />
          <Check
            label="Markers"
            checked={spec.markers}
            onChange={(markers) => set((s) => ({ ...s, markers }), "Markers")}
          />
        </Section>
      )
    }
    return null
  }
  if (part === "chartArea" && chartHasAxes(spec.type)) {
    return (
      <Section title="Chart Options">
        <p className="cfp-note">
          Click a part of the chart, or pick it in the list above, to format it. Chart Design › Add Chart Element shows
          the parts that are hidden.
        </p>
      </Section>
    )
  }
  return null
}

export function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <details className="cfp-section" open>
      <summary>{title}</summary>
      <div className="cfp-section-body">{children}</div>
    </details>
  )
}

export function Row({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="cfp-row">
      <span>{label}</span>
      {children}
    </div>
  )
}

export function Check({
  label,
  checked,
  onChange,
}: {
  label: string
  checked: boolean
  onChange: (on: boolean) => void
}) {
  return (
    <label className="dialog-radio">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

export function Choice({
  name,
  value,
  options,
  onChange,
}: {
  name: string
  value: string
  options: [string, string][]
  onChange: (value: string) => void
}) {
  return (
    <div className="cfp-choices" role="radiogroup">
      {options.map(([v, label]) => (
        <label key={v} className="dialog-radio">
          <input type="radio" name={`cfp-${name}`} checked={value === v} onChange={() => onChange(v)} />
          <span>{label}</span>
        </label>
      ))}
    </div>
  )
}

/** A colour swatch that opens the theme colours; `automatic` offers the style's own colour back. */
export function ColorButton({
  label,
  color,
  automatic = false,
  theme,
  onPick,
}: {
  label: string
  color: string | undefined
  automatic?: boolean
  theme: SheetController["wb"]["theme"]
  onPick: (color: string | null) => void
}) {
  const [open, setOpen] = useState(false)
  return (
    <Popover
      open={open}
      onOpenChange={setOpen}
      trigger={
        <button type="button" className="dialog-field cfp-color" aria-label={label}>
          <span style={{ background: color && color !== "none" ? color : "transparent" }} />
          {color ? "" : "Automatic"}
        </button>
      }
    >
      <ColorPanel
        noneLabel={automatic ? "Automatic" : undefined}
        current={color}
        theme={theme}
        onPick={(hex) => {
          setOpen(false)
          onPick(hex ? hex.toLowerCase() : null)
        }}
      />
    </Popover>
  )
}

/** A number typed and kept on Enter or leaving the box (one undo step, not one per key); empty is "Auto". */
export function NumberInput({
  label,
  value,
  unit,
  placeholder,
  min,
  max,
  step,
  onCommit,
}: {
  label: string
  value: number | undefined
  unit?: string
  placeholder?: string
  min?: number
  max?: number
  step?: number
  onCommit: (value: number | undefined) => void
}) {
  const shown = value === undefined ? "" : String(value)
  const [text, setText] = useState(shown)
  useEffect(() => setText(shown), [shown])
  const commit = () => {
    const raw = text.trim()
    if (raw === shown) return
    if (raw === "") return onCommit(undefined)
    let n = Number(raw.replace(",", "."))
    if (!Number.isFinite(n)) return setText(shown)
    if (min !== undefined) n = Math.max(min, n)
    if (max !== undefined) n = Math.min(max, n)
    onCommit(n)
  }
  return (
    <span className="cfp-number">
      <input
        className="dialog-field"
        aria-label={label}
        inputMode="decimal"
        type="number"
        step={step}
        min={min}
        max={max}
        value={text}
        placeholder={placeholder}
        onChange={(e) => setText(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit()
        }}
      />
      {unit ? <span className="cfp-unit">{unit}</span> : null}
    </span>
  )
}

function TextInput({ label, value, onCommit }: { label: string; value: string; onCommit: (text: string) => void }) {
  const [text, setText] = useState(value)
  useEffect(() => setText(value), [value])
  const commit = () => {
    if (text !== value) onCommit(text)
  }
  return (
    <input
      className="dialog-field"
      aria-label={label}
      value={text}
      onChange={(e) => setText(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") commit()
      }}
    />
  )
}
