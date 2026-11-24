import { memo, useState } from "react"
import { ContextMenu, type MenuItem } from "@/shell/ContextMenu"
import {
  TABLE_STYLE_GROUPS,
  tableCellLook,
  tableStyleDef,
  type Border,
  type CustomTableStyle,
  type Table,
  type TableStyleDef,
  type Theme,
} from "@cells/core"

/** The made-up table each thumbnail draws: a header and four rows of five columns, as Excel's gallery shows. */
const SAMPLE: Table = {
  id: 0,
  name: "Sample",
  range: { r0: 0, c0: 0, r1: 4, c1: 4 },
  header: true,
  totals: false,
  columns: ["A", "B", "C", "D", "E"].map((name) => ({ name })),
  style: null,
  firstColumn: false,
  lastColumn: false,
  rowStripes: true,
  columnStripes: false,
  autoFilter: null,
}

const CW = 12
const RH = 8
const width = (b: Border) => (b.style === "thick" || b.style === "double" ? 1.6 : b.style === "medium" ? 1.3 : 0.8)

const OPTION_KEYS = ["header", "totals", "firstColumn", "lastColumn", "rowStripes", "columnStripes"] as const

/** A miniature of a table style (paper: it stays light in dark mode). */
export const TableStyleThumb = memo(function TableStyleThumb({
  style,
  options,
  theme,
  custom,
  def: given,
  scale = 1,
}: {
  /** The workbook's theme, whose colours the built-in styles take. */
  theme?: Theme
  style: string | null
  options?: Partial<Pick<Table, "header" | "totals" | "firstColumn" | "lastColumn" | "rowStripes" | "columnStripes">>
  /** The workbook's own styles (the gallery's Custom group). */
  custom?: readonly CustomTableStyle[]
  /** A style being made (New Table Style's preview), in place of `style`'s. */
  def?: TableStyleDef | null
  /** Drawn larger (the dialog's preview). */
  scale?: number
}) {
  // Only the style options: a real table passed as them would bring its own range along.
  const picked = Object.fromEntries(OPTION_KEYS.filter((k) => options?.[k] !== undefined).map((k) => [k, options![k]]))
  const t: Table = { ...SAMPLE, ...picked, style }
  const def = given !== undefined ? given : tableStyleDef(style, theme, custom)
  const rows = t.range.r1 - t.range.r0 + 1
  const cols = 5
  const fills: React.ReactNode[] = []
  const marks: React.ReactNode[] = []
  const lines: React.ReactNode[] = []
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const look = def ? tableCellLook(t, r, c, def) : null
      const x = c * CW + 0.5
      const y = r * RH + 0.5
      if (look?.fill) fills.push(<rect key={`f${r}-${c}`} x={x} y={y} width={CW} height={RH} fill={look.fill} />)
      marks.push(
        <line
          key={`m${r}-${c}`}
          x1={x + 3}
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
  if (!def)
    // "None": the gridlines only.
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
      width={(cols * CW + 1) * scale}
      height={(rows * RH + 1) * scale}
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

/** "TableStyleMedium2" → "Medium Blue, Table Style Medium 2" is Excel's tooltip; the number part is enough here. */
export function tableStyleLabel(name: string | null): string {
  if (!name) return "None"
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name)
  return m ? `${m[1]} ${m[2]}` : name
}

/**
 * The styles gallery of Format as Table and Table Design: Light, Medium and Dark, as miniatures. `withNone` puts
 * Excel's "None" first in Light (Table Design only). `options` draws the thumbnails with the table's own options.
 */
export function TableStyleGallery({
  current,
  withNone = false,
  options,
  theme,
  onPick,
  custom = [],
  menuFor,
}: {
  theme?: Theme
  current?: string | null
  withNone?: boolean
  options?: Parameters<typeof TableStyleThumb>[0]["options"]
  onPick: (style: string | null) => void
  /** The workbook's own styles, shown first as Custom. */
  custom?: readonly CustomTableStyle[]
  /** A style's right-click menu (Apply, Modify, Duplicate, Delete, Set As Default). */
  menuFor?: (style: string | null) => MenuItem[]
}) {
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null)
  const groups = custom.length
    ? [{ label: "Custom", names: custom.map((s) => s.name) }, ...TABLE_STYLE_GROUPS]
    : TABLE_STYLE_GROUPS
  return (
    <div className="table-style-gallery" role="menu">
      {menu ? <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} /> : null}
      {groups.map((group) => (
        <div key={group.label}>
          <div className="color-heading">{group.label}</div>
          <div className="table-style-grid">
            {(group.label === "Light" && withNone ? [null, ...group.names] : group.names).map((name) => (
              <button
                key={name ?? "none"}
                type="button"
                role="menuitem"
                className={`table-style-item${current !== undefined && current === name ? " is-current" : ""}`}
                title={tableStyleLabel(name)}
                aria-label={tableStyleLabel(name)}
                data-keep-light
                onClick={() => onPick(name)}
                onContextMenu={(e) => {
                  if (!menuFor) return
                  e.preventDefault()
                  // The style's menu, not the ribbon's (Add to Quick Access Toolbar…).
                  e.stopPropagation()
                  setMenu({ x: e.clientX, y: e.clientY, items: menuFor(name) })
                }}
              >
                <TableStyleThumb style={name} options={options} theme={theme} custom={custom} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}
