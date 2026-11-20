import { memo } from "react"
import {
  newPivotTable,
  PIVOT_STYLE_GROUPS,
  pivotCellLook,
  pivotLayoutOf,
  pivotStyleDef,
  withFieldIn,
  type Border,
  type PivotCache,
  type PivotTable,
  type PivotTableStyle,
  type Theme,
} from "@cells/core"

/** The made-up data the thumbnails lay out: two regions of two products, by three channels, with totals. */
const SAMPLE_CACHE: PivotCache = {
  id: 0,
  source: { kind: "range", sheet: "Sample", ref: "A1:D2" },
  fields: [
    { name: "Region", items: ["A", "B"] },
    { name: "Product", items: ["x", "y"] },
    { name: "Channel", items: ["p", "q", "r"] },
    { name: "Units", items: null },
  ],
  records: ["A", "B"].flatMap((r) => ["x", "y"].flatMap((p) => ["p", "q", "r"].map((c) => [r, p, c, 1]))),
  refreshOnLoad: false,
}

const BASE = (() => {
  let t = newPivotTable("Sample", SAMPLE_CACHE, { row: 0, col: 0 })
  for (const [f, area] of [
    [0, "rows"],
    [1, "rows"],
    [2, "cols"],
    [3, "values"],
  ] as const)
    t = withFieldIn(t, SAMPLE_CACHE, f, area)
  return t
})()

/** One sample table per set of style options (the layout is kept per table object). */
const samples = new Map<string, PivotTable>()
function sample(style: string | null, options: Partial<PivotTableStyle>): PivotTable {
  const key = JSON.stringify([style, options])
  let t = samples.get(key)
  if (!t) samples.set(key, (t = { ...BASE, style: { ...BASE.style, ...options, name: style } }))
  return t
}

const CW = 12
const RH = 6
const width = (b: Border) => (b.style === "thick" || b.style === "double" ? 1.6 : b.style === "medium" ? 1.3 : 0.8)

/** A miniature of a PivotTable style (paper: it stays light in dark mode). */
export const PivotStyleThumb = memo(function PivotStyleThumb({
  style,
  options = {},
  theme,
}: {
  style: string | null
  options?: Partial<PivotTableStyle>
  theme?: Theme
}) {
  const t = sample(style, options)
  const range = pivotLayoutOf(t, SAMPLE_CACHE).range
  const rows = range.r1 - range.r0 + 1
  const cols = range.c1 - range.c0 + 1
  const def = pivotStyleDef(style, theme)
  const fills: React.ReactNode[] = []
  const marks: React.ReactNode[] = []
  const lines: React.ReactNode[] = []
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const look = def ? pivotCellLook(t, SAMPLE_CACHE, range.r0 + r, range.c0 + c, def) : null
      const x = c * CW + 0.5
      const y = r * RH + 0.5
      if (look?.fill) fills.push(<rect key={`f${r}-${c}`} x={x} y={y} width={CW} height={RH} fill={look.fill} />)
      const indent = (look?.indent ?? 0) * 2
      marks.push(
        <line
          key={`m${r}-${c}`}
          x1={x + 3 + indent}
          x2={x + CW - 3}
          y1={y + RH / 2}
          y2={y + RH / 2}
          stroke={look?.fontColor ?? "#404040"}
          strokeWidth={look?.bold ? 1.6 : 1}
          opacity={0.75}
        />,
      )
      const edge = (b: Border | undefined, k: string, x1: number, y1: number, x2: number, y2: number) => {
        if (b)
          lines.push(<line key={k} x1={x1} y1={y1} x2={x2} y2={y2} stroke={b.color ?? "#000"} strokeWidth={width(b)} />)
      }
      edge(look?.borderTop, `t${r}-${c}`, x, y, x + CW, y)
      edge(look?.borderBottom, `b${r}-${c}`, x, y + RH, x + CW, y + RH)
      edge(look?.borderLeft, `l${r}-${c}`, x, y, x, y + RH)
      edge(look?.borderRight, `r${r}-${c}`, x + CW, y, x + CW, y + RH)
    }
  if (!style)
    for (let r = 0; r <= rows; r++)
      lines.push(
        <line
          key={`g${r}`}
          x1={0.5}
          x2={cols * CW + 0.5}
          y1={r * RH + 0.5}
          y2={r * RH + 0.5}
          stroke="#d0d0d0"
          strokeWidth={0.6}
        />,
      )
  return (
    <svg
      className="table-style-thumb"
      width={cols * CW + 1}
      height={rows * RH + 1}
      viewBox={`0 0 ${cols * CW + 1} ${rows * RH + 1}`}
      aria-hidden
    >
      <rect x={0} y={0} width={cols * CW + 1} height={rows * RH + 1} fill="#fff" />
      {fills}
      {marks}
      {lines}
    </svg>
  )
})

/** "PivotStyleLight16" → "Light 16". */
export function pivotStyleLabel(name: string | null): string {
  if (!name) return "None"
  const m = /^PivotStyle(Light|Medium|Dark)(\d+)$/.exec(name)
  return m ? `${m[1]} ${m[2]}` : name
}

/** Design › PivotTable Styles: Light (with None first), Medium and Dark, drawn with the table's style options. */
export function PivotStyleGallery({
  current,
  options,
  theme,
  onPick,
}: {
  current: string | null
  options?: Partial<PivotTableStyle>
  theme?: Theme
  onPick: (style: string | null) => void
}) {
  return (
    <div className="table-style-gallery pivot-style-gallery" role="menu" aria-label="PivotTable Styles">
      {PIVOT_STYLE_GROUPS.map((group) => (
        <div key={group.label}>
          <div className="color-heading">{group.label}</div>
          <div className="table-style-grid">
            {(group.label === "Light" ? [null, ...group.names] : group.names).map((name) => (
              <button
                key={name ?? "none"}
                type="button"
                role="menuitem"
                aria-label={pivotStyleLabel(name)}
                title={pivotStyleLabel(name)}
                className={`table-style-item${name === current ? " is-current" : ""}`}
                data-keep-light
                onClick={() => onPick(name)}
              >
                <PivotStyleThumb style={name} options={options} theme={theme} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}
