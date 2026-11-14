import {
  DATE_ORDERS,
  decodeText,
  defaultImport,
  detectBreaks,
  splitText,
  textLines,
  TEXT_ENCODINGS,
  type ColumnFormat,
  type TextEncoding,
  type TextImport,
} from "@cells/io"
import { useMemo, useRef, useState } from "react"
import { cn } from "@/lib/cn"
import { Dialog } from "@/shell/Dialog"

const PREVIEW_ROWS = 100

const DELIMITERS = [
  { id: "tab", char: "\t", label: "Tab" },
  { id: "semicolon", char: ";", label: "Semicolon" },
  { id: "comma", char: ",", label: "Comma" },
  { id: "space", char: " ", label: "Space" },
] as const

/**
 * Excel's Text Import Wizard, in its three steps: how the data is laid out (delimited or fixed width, the first row,
 * the file's encoding), where the columns split, and each column's format. Finish hands back the choices.
 */
export function TextImportWizard({
  bytes,
  fileName,
  onFinish,
  onCancel,
}: {
  bytes: Uint8Array
  fileName: string
  onFinish: (choices: TextImport) => void
  onCancel: () => void
}) {
  const first = useMemo(() => decodeText(bytes), [bytes])
  const [step, setStep] = useState(1)
  const [encoding, setEncoding] = useState<TextEncoding>(first.encoding)
  const text = useMemo(
    () => (encoding === first.encoding ? first.text : decodeText(bytes, encoding).text),
    [bytes, encoding, first],
  )
  const initial = useMemo(() => defaultImport(first.text, fileName), [first, fileName])
  const [kind, setKind] = useState<TextImport["kind"]>(() => (looksFixed(first.text) ? "fixed" : "delimited"))
  const [startRow, setStartRow] = useState(1)
  const [delimiters, setDelimiters] = useState<string[]>(initial.delimiters)
  const [other, setOther] = useState("")
  const [consecutive, setConsecutive] = useState(false)
  const [qualifier, setQualifier] = useState<string | null>('"')
  const lines = useMemo(() => textLines(text, PREVIEW_ROWS + 1000), [text])
  const [breaks, setBreaks] = useState<number[]>(() => detectBreaks(textLines(first.text, 200)))
  const [columns, setColumns] = useState<ColumnFormat[]>([])
  const [selected, setSelected] = useState(0)
  const [advanced, setAdvanced] = useState(false)
  const [decimal, setDecimal] = useState(".")
  const [thousands, setThousands] = useState(",")

  const choices: TextImport = {
    encoding,
    kind,
    delimiters: [...delimiters, ...(other ? [other[0]] : [])],
    qualifier,
    consecutive,
    breaks,
    startRow,
    columns,
    decimal,
    thousands,
  }
  const preview = useMemo(
    () => splitText(text, { ...choices, columns: [] }, PREVIEW_ROWS),
    [text, kind, startRow, delimiters.join(""), other, consecutive, qualifier, breaks.join(",")],
  )
  const width = Math.max(1, ...preview.map((r) => r.length))
  const formatOf = (col: number) => columns[col] ?? "general"
  const setFormat = (format: ColumnFormat) => {
    const next = [...columns]
    while (next.length <= selected) next.push("general")
    next[selected] = format
    setColumns(next)
  }

  const next = () => (step < 3 ? setStep(step + 1) : onFinish(choices))
  const toggle = (char: string, on: boolean) =>
    setDelimiters((list) => (on ? [...list.filter((c) => c !== char), char] : list.filter((c) => c !== char)))
  const dateFormat = DATE_ORDERS.find((o) => o === formatOf(selected)) ?? "MDY"

  return (
    <Dialog
      title={`Text Import Wizard - Step ${step} of 3`}
      width={640}
      onCancel={onCancel}
      onAccept={next}
      actions={
        <>
          <button type="button" className="dialog-btn" onClick={onCancel}>
            Cancel
          </button>
          <button type="button" className="dialog-btn" disabled={step === 1} onClick={() => setStep(step - 1)}>
            &lt; Back
          </button>
          <button type="button" className="dialog-btn" disabled={step === 3} onClick={() => setStep(step + 1)}>
            Next &gt;
          </button>
          <button type="button" className="dialog-btn is-primary" onClick={() => onFinish(choices)}>
            Finish
          </button>
        </>
      }
    >
      <div className="tiw">
        {step === 1 ? (
          <>
            <p className="dialog-message">
              The Text Wizard has determined that your data is {kind === "fixed" ? "Fixed Width" : "Delimited"}. If this
              is correct, choose Next, or choose the data type that best describes your data.
            </p>
            <fieldset className="tiw-fieldset">
              <legend>Original data type</legend>
              <label className="dialog-radio">
                <input type="radio" checked={kind === "delimited"} onChange={() => setKind("delimited")} />
                <span className="tiw-choice">Delimited</span>
                <span className="tiw-hint">Characters such as commas or tabs separate each field.</span>
              </label>
              <label className="dialog-radio">
                <input type="radio" checked={kind === "fixed"} onChange={() => setKind("fixed")} />
                <span className="tiw-choice">Fixed width</span>
                <span className="tiw-hint">Fields are aligned in columns with spaces between each field.</span>
              </label>
            </fieldset>
            <div className="tiw-row">
              <label className="dialog-label">
                Start import at row:{" "}
                <input
                  className="dialog-field tiw-number"
                  type="number"
                  min={1}
                  max={Math.max(1, lines.length)}
                  value={startRow}
                  onChange={(e) => setStartRow(Math.max(1, Math.floor(Number(e.target.value)) || 1))}
                />
              </label>
              <label className="dialog-label">
                File origin:{" "}
                <select
                  className="dialog-field"
                  value={encoding}
                  onChange={(e) => setEncoding(e.target.value as TextEncoding)}
                >
                  {TEXT_ENCODINGS.map((e) => (
                    <option key={e.id} value={e.id}>
                      {e.label}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <p className="tiw-caption">Preview of file {fileName}.</p>
            <pre className="tiw-lines" aria-label="Preview">
              {lines.slice(0, PREVIEW_ROWS).map((line, i) => (
                <div key={i} className={cn(i + 1 < startRow && "is-skipped")}>
                  <span className="tiw-line-no">{i + 1}</span>
                  {line}
                </div>
              ))}
            </pre>
          </>
        ) : null}

        {step === 2 && kind === "delimited" ? (
          <>
            <p className="dialog-message">
              This screen lets you set the delimiters your data contains. You can see how your text is affected in the
              preview below.
            </p>
            <div className="tiw-columns">
              <fieldset className="tiw-fieldset">
                <legend>Delimiters</legend>
                {DELIMITERS.map((d) => (
                  <label key={d.id} className="dialog-radio">
                    <input
                      type="checkbox"
                      checked={delimiters.includes(d.char)}
                      onChange={(e) => toggle(d.char, e.target.checked)}
                    />
                    {d.label}
                  </label>
                ))}
                <label className="dialog-radio">
                  <input type="checkbox" checked={other !== ""} onChange={(e) => !e.target.checked && setOther("")} />
                  Other:
                  <input
                    className="dialog-field tiw-other"
                    maxLength={1}
                    aria-label="Other delimiter"
                    value={other}
                    onChange={(e) => setOther(e.target.value.slice(0, 1))}
                  />
                </label>
              </fieldset>
              <div className="tiw-stack">
                <label className="dialog-radio">
                  <input type="checkbox" checked={consecutive} onChange={(e) => setConsecutive(e.target.checked)} />
                  Treat consecutive delimiters as one
                </label>
                <label className="dialog-label">
                  Text qualifier:{" "}
                  <select
                    className="dialog-field"
                    value={qualifier ?? ""}
                    onChange={(e) => setQualifier(e.target.value || null)}
                  >
                    <option value={'"'}>"</option>
                    <option value="'">'</option>
                    <option value="">{"{none}"}</option>
                  </select>
                </label>
              </div>
            </div>
            <p className="tiw-caption">Data preview</p>
            <PreviewTable rows={preview} width={width} />
          </>
        ) : null}

        {step === 2 && kind === "fixed" ? (
          <>
            <p className="dialog-message">
              This screen lets you set field widths (column breaks). Lines with arrows signify a column break.
            </p>
            <ul className="tiw-help">
              <li>To CREATE a break line, click at the desired position.</li>
              <li>To DELETE a break line, double click on the line.</li>
            </ul>
            <p className="tiw-caption">Data preview</p>
            <FixedPreview
              lines={lines.slice(startRow - 1, startRow - 1 + PREVIEW_ROWS)}
              breaks={breaks}
              onChange={setBreaks}
            />
          </>
        ) : null}

        {step === 3 ? (
          <>
            <p className="dialog-message">This screen lets you select each column and set the Data Format.</p>
            <div className="tiw-columns">
              <fieldset className="tiw-fieldset">
                <legend>Column data format</legend>
                <label className="dialog-radio">
                  <input
                    type="radio"
                    checked={formatOf(selected) === "general"}
                    onChange={() => setFormat("general")}
                  />
                  General
                </label>
                <label className="dialog-radio">
                  <input type="radio" checked={formatOf(selected) === "text"} onChange={() => setFormat("text")} />
                  Text
                </label>
                <label className="dialog-radio">
                  <input
                    type="radio"
                    checked={(DATE_ORDERS as readonly string[]).includes(formatOf(selected))}
                    onChange={() => setFormat(dateFormat)}
                  />
                  Date:
                  <select
                    className="dialog-field"
                    aria-label="Date order"
                    value={dateFormat}
                    onChange={(e) => setFormat(e.target.value as ColumnFormat)}
                  >
                    {DATE_ORDERS.map((o) => (
                      <option key={o} value={o}>
                        {o}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="dialog-radio">
                  <input type="radio" checked={formatOf(selected) === "skip"} onChange={() => setFormat("skip")} />
                  Do not import column (skip)
                </label>
              </fieldset>
              <div className="tiw-stack">
                <p className="tiw-hint">
                  'General' converts numeric values to numbers, date values to dates, and all remaining values to text.
                </p>
                <button type="button" className="dialog-btn" onClick={() => setAdvanced(!advanced)}>
                  Advanced...
                </button>
                {advanced ? (
                  <div className="tiw-advanced">
                    <label className="dialog-label">
                      Decimal separator:{" "}
                      <select className="dialog-field" value={decimal} onChange={(e) => setDecimal(e.target.value)}>
                        {[".", ","].map((c) => (
                          <option key={c} value={c}>
                            {c}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className="dialog-label">
                      Thousands separator:{" "}
                      <select className="dialog-field" value={thousands} onChange={(e) => setThousands(e.target.value)}>
                        {[",", ".", " ", "'"].map((c) => (
                          <option key={c} value={c}>
                            {c === " " ? "(space)" : c}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                ) : null}
              </div>
            </div>
            <p className="tiw-caption">Data preview</p>
            <PreviewTable
              rows={preview}
              width={width}
              headers={Array.from({ length: width }, (_, c) => FORMAT_LABEL[formatOf(c)] ?? formatOf(c))}
              selected={selected}
              onSelect={setSelected}
              skipped={(c) => formatOf(c) === "skip"}
            />
          </>
        ) : null}
      </div>
    </Dialog>
  )
}

const FORMAT_LABEL: Partial<Record<ColumnFormat, string>> = {
  general: "General",
  text: "Text",
  skip: "Skip Column",
}

function PreviewTable({
  rows,
  width,
  headers,
  selected,
  onSelect,
  skipped,
}: {
  rows: string[][]
  width: number
  headers?: string[]
  selected?: number
  onSelect?: (col: number) => void
  skipped?: (col: number) => boolean
}) {
  const cols = Array.from({ length: width }, (_, c) => c)
  return (
    <div className="tiw-preview">
      <table>
        {headers ? (
          <thead>
            <tr>
              {cols.map((c) => (
                <th
                  key={c}
                  className={cn(c === selected && "is-selected")}
                  onClick={() => onSelect?.(c)}
                  aria-selected={c === selected}
                >
                  {headers[c]}
                </th>
              ))}
            </tr>
          </thead>
        ) : null}
        <tbody>
          {rows.map((row, r) => (
            <tr key={r}>
              {cols.map((c) => (
                <td
                  key={c}
                  className={cn(c === selected && "is-selected", skipped?.(c) && "is-skipped")}
                  onClick={() => onSelect?.(c)}
                >
                  {row[c] ?? ""}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

/** The fixed-width preview: a character ruler, the lines, and the column breaks to add, move and remove. */
function FixedPreview({
  lines,
  breaks,
  onChange,
}: {
  lines: string[]
  breaks: number[]
  onChange: (breaks: number[]) => void
}) {
  const ref = useRef<HTMLDivElement>(null)
  const longest = Math.max(10, ...lines.map((l) => l.length)) + 2
  const columnAt = (clientX: number) => {
    const box = ref.current
    if (!box) return 0
    const ch = box.querySelector<HTMLElement>(".tiw-measure")?.getBoundingClientRect().width ?? 7
    return Math.max(1, Math.round((clientX - box.getBoundingClientRect().left) / (ch / 10)))
  }
  const ruler = Array.from({ length: Math.ceil(longest / 10) }, (_, i) => String((i + 1) * 10).padStart(10, " "))
  return (
    <div className="tiw-fixed">
      <div
        ref={ref}
        className="tiw-fixed-body"
        onClick={(e) => {
          const col = columnAt(e.clientX)
          if (!breaks.includes(col) && col < longest) onChange([...breaks, col].sort((a, b) => a - b))
        }}
      >
        <span className="tiw-measure" aria-hidden>
          0000000000
        </span>
        <div className="tiw-ruler" aria-hidden>
          {ruler.join("")}
        </div>
        {lines.map((line, i) => (
          <div key={i}>{line || " "}</div>
        ))}
        {breaks.map((b) => (
          <span
            key={b}
            className="tiw-break"
            style={{ left: `${b}ch` }}
            role="separator"
            aria-label={`Column break at ${b}`}
            onClick={(e) => e.stopPropagation()}
            onDoubleClick={(e) => {
              e.stopPropagation()
              onChange(breaks.filter((x) => x !== b))
            }}
          />
        ))}
      </div>
    </div>
  )
}

/** Whether a text looks like fixed-width columns: no usual delimiter, and runs of spaces lined up across lines. */
function looksFixed(text: string): boolean {
  const lines = textLines(text, 30).filter((l) => l.trim())
  if (lines.length < 2) return false
  if (lines.some((l) => /[,;\t|]/.test(l))) return false
  return detectBreaks(lines).length > 0 && lines.every((l) => / {2,}/.test(l))
}
