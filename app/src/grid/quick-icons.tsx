import type { ReactNode } from "react"
import { ChartTypeIcon } from "@/ribbon/chart-commands"

/**
 * Quick Analysis's pictures, one per choice: drawn on one 32 px grid with one stroke, in the ribbon's icon colours
 * (`--icon-base`, `--icon-accent`, `--icon-muted`), so they theme with the app. Charts use the ribbon's chart
 * gallery icons.
 */

const base = { stroke: "var(--icon-base)", fill: "none", strokeWidth: 1.25 } as const
const muted = { stroke: "var(--icon-muted)", fill: "none", strokeWidth: 1 } as const
const accent = "var(--icon-accent)"
const soft = "var(--icon-accent-soft)"

/** A small sheet: 3 columns × 4 rows of cells from (5,5) to (27,27). */
function Sheet({ children, rows = 4 }: { children?: ReactNode; rows?: number }) {
  const h = 22 / rows
  return (
    <>
      {children}
      {Array.from({ length: rows - 1 }, (_, i) => (
        <path key={`r${i}`} d={`M5 ${5 + h * (i + 1)}h22`} style={muted} />
      ))}
      <path d="M12.33 5v22M19.67 5v22" style={muted} />
      <rect x={5} y={5} width={22} height={22} rx={1} style={base} />
    </>
  )
}

const cell = (col: number, row: number, fill: string, rows = 4) => {
  const h = 22 / rows
  return <rect key={`${col}:${row}:${fill}`} x={5 + col * 7.33} y={5 + row * h} width={7.33} height={h} fill={fill} />
}

function Svg({ children }: { children: ReactNode }) {
  return (
    <svg width={32} height={32} viewBox="0 0 32 32" aria-hidden className="qa-icon">
      {children}
    </svg>
  )
}

const GREEN = "#63BE7B"
const YELLOW = "#FFD966"
const RED = "#F8696B"
const BAR = "#638EC6"

function formattingIcon(id: string): ReactNode {
  switch (id) {
    case "cf:dataBar":
      return (
        <Sheet>
          {[16, 9, 19, 5].map((w, i) => (
            <rect key={i} x={6} y={6.5 + i * 5.5} width={w} height={3} rx={0.5} fill={BAR} />
          ))}
        </Sheet>
      )
    case "cf:colorScale":
      return <Sheet>{[GREEN, YELLOW, "#FFA36B", RED].flatMap((c, r) => [0, 1, 2].map((col) => cell(col, r, c)))}</Sheet>
    case "cf:iconSet":
      return (
        <>
          <Sheet />
          {[
            { y: 8.5, d: "M-3 2.5 0 -2.5 3 2.5z", c: GREEN },
            { y: 16, d: "M-2.5 -3 2.5 0 -2.5 3z", c: "#E5B800" },
            { y: 23.5, d: "M-3 -2.5 0 2.5 3 -2.5z", c: RED },
          ].map((a) => (
            <path key={a.y} d={a.d} transform={`translate(8.7 ${a.y})`} fill={a.c} />
          ))}
        </>
      )
    case "cf:greaterThan":
    case "cf:lessThan":
      return (
        <>
          <Sheet>{[cell(1, 0, soft), cell(1, 2, soft), cell(2, 1, soft)]}</Sheet>
          <path
            d={id === "cf:greaterThan" ? "M20 18.5l7 3.5-7 3.5" : "M27 18.5l-7 3.5 7 3.5"}
            style={{ ...base, stroke: accent, strokeWidth: 2, strokeLinejoin: "round", strokeLinecap: "round" }}
          />
        </>
      )
    case "cf:equal":
      return (
        <>
          <Sheet>{[cell(0, 1, soft), cell(2, 3, soft)]}</Sheet>
          <path d="M19.5 20h7M19.5 24h7" style={{ ...base, stroke: accent, strokeWidth: 2, strokeLinecap: "round" }} />
        </>
      )
    case "cf:top10":
      return (
        <>
          <Sheet>{[0, 1, 2].map((c) => cell(c, 0, soft))}</Sheet>
          <path d="M23 25v-7m-3 3 3-3 3 3" style={{ ...base, stroke: accent, strokeWidth: 1.75 }} />
        </>
      )
    case "cf:text":
      return (
        <Sheet>
          {cell(0, 1, soft)}
          {[0, 1, 2, 3].map((r) => (
            <path key={r} d={`M6.5 ${7.75 + r * 5.5}h${r === 1 ? 4.5 : 3.5}`} style={{ ...base, strokeWidth: 1 }} />
          ))}
        </Sheet>
      )
    case "cf:duplicate":
      return <Sheet>{[cell(0, 0, soft), cell(0, 2, soft), cell(2, 1, soft), cell(2, 3, soft)]}</Sheet>
    case "cf:unique":
      return <Sheet>{[cell(1, 1, soft), cell(0, 3, soft)]}</Sheet>
    case "cf:lastMonth":
    case "cf:lastWeek":
      return (
        <>
          <rect x={5} y={7} width={22} height={20} rx={1.5} style={base} />
          <path d="M5 12h22M10 5v4M22 5v4" style={base} />
          {id === "cf:lastWeek" ? (
            <rect x={7} y={17} width={18} height={3.5} fill={accent} />
          ) : (
            <rect x={7} y={14} width={18} height={11} fill={soft} />
          )}
        </>
      )
    case "cf:clear":
      return (
        <>
          <Sheet>{[cell(0, 0, soft), cell(1, 1, soft)]}</Sheet>
          <path
            d="M18 26l6-6 3.5 3.5-3 3z"
            style={{ ...base, fill: "var(--icon-danger-soft, #fde7e9)", stroke: "var(--icon-danger, #c50f1f)" }}
          />
        </>
      )
  }
  return <Sheet />
}

/** Totals: the row below (Excel draws it blue) or the column right (yellow), with the function's sign. */
function totalIcon(side: string, kind: string): ReactNode {
  const below = side === "below"
  const band = below ? <rect x={5} y={21.5} width={22} height={5.5} fill="#9DC3E6" /> : null
  const column = below ? null : <rect x={19.67} y={5} width={7.33} height={22} fill="#FFE699" />
  const sign: Record<string, ReactNode> = {
    sum: <path d="M14.5 10.5h-5l3 4.5-3 4.5h5" style={{ ...base, strokeWidth: 1.5 }} />,
    average: (
      <>
        <circle cx={12} cy={15} r={3.5} style={{ ...base, strokeWidth: 1.5 }} />
        <path d="M9 18.5l6-7" style={{ ...base, strokeWidth: 1.5 }} />
      </>
    ),
    count: <path d="M10.5 10v9.5M14 10v9.5M8.5 13h7.5M8 16.5h7.5" style={{ ...base, strokeWidth: 1.25 }} />,
    percent: (
      <>
        <circle cx={9.5} cy={11.5} r={1.5} style={base} />
        <circle cx={14.5} cy={18} r={1.5} style={base} />
        <path d="M15 10.5l-6 8.5" style={base} />
      </>
    ),
    running: <path d="M8 19l3-3 2.5 1.5L16 11m-2.5 0H16v2.5" style={{ ...base, strokeWidth: 1.5 }} />,
  }
  return (
    <>
      {band}
      {column}
      <rect x={5} y={5} width={22} height={22} rx={1} style={base} />
      <path d={below ? "M5 21.5h22" : "M19.67 5v22"} style={muted} />
      <rect x={6.5} y={7.5} width={10.5} height={13.5} rx={1.5} fill="var(--app-popover-bg)" />
      {sign[kind]}
    </>
  )
}

function tableIcon(pivot: boolean): ReactNode {
  if (pivot)
    return (
      <>
        <Sheet>{[cell(0, 0, soft), cell(1, 0, soft), cell(2, 0, soft), cell(0, 1, soft), cell(0, 2, soft)]}</Sheet>
        <path d="M17 22h7m-2-2 2 2-2 2" style={{ ...base, stroke: accent }} />
      </>
    )
  return (
    <Sheet>
      {[0, 1, 2].map((c) => cell(c, 0, accent))}
      {[0, 1, 2].map((c) => cell(c, 2, soft))}
    </Sheet>
  )
}

function sparkIcon(type: string): ReactNode {
  const frame = <rect x={4} y={9} width={24} height={14} rx={1} style={muted} />
  if (type === "line")
    return (
      <>
        {frame}
        <path
          d="M6 19l4-5 4 3 4-6 4 4 4-3"
          style={{ ...base, stroke: BAR, strokeWidth: 1.5, strokeLinejoin: "round" }}
        />
      </>
    )
  const heights = type === "column" ? [5, 8, 4, 10, 7, 9] : [4, 4, -4, 4, -4, 4]
  return (
    <>
      {frame}
      {heights.map((h, i) => {
        const x = 6.5 + i * 3.6
        if (type === "column") return <rect key={i} x={x} y={21 - h} width={2.4} height={h} fill={BAR} />
        return <rect key={i} x={x} y={h > 0 ? 12 : 16} width={2.4} height={4} fill={h > 0 ? BAR : RED} />
      })}
    </>
  )
}

/** The picture for a Quick Analysis choice. */
export function QuickIcon({ id }: { id: string }) {
  const [what, a, b] = id.split(":")
  if (what === "chart") return <ChartTypeIcon id={a} size={32} />
  if (what === "more")
    return (
      <Svg>
        <rect x={5} y={7} width={22} height={18} rx={1.5} style={base} />
        <path d="M10 21v-4m4 4v-8m4 8v-6" style={{ ...base, stroke: accent, strokeWidth: 2 }} />
        <circle cx={24} cy={11} r={0.9} fill="var(--icon-base)" />
        <circle cx={24} cy={14} r={0.9} fill="var(--icon-base)" />
        <circle cx={24} cy={17} r={0.9} fill="var(--icon-base)" />
      </Svg>
    )
  if (what === "total") return <Svg>{totalIcon(a, b)}</Svg>
  if (what === "table" || what === "pivot") return <Svg>{tableIcon(what === "pivot")}</Svg>
  if (what === "spark") return <Svg>{sparkIcon(a)}</Svg>
  return <Svg>{formattingIcon(id)}</Svg>
}

/** The button at the selection's corner: a small sheet with a lightning bolt, as Excel's. */
export function QuickButtonIcon({ size }: { size: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" aria-hidden>
      <rect x={1.5} y={2.5} width={10} height={9} rx={0.75} style={{ ...base, strokeWidth: 1 }} />
      <path d="M1.5 5.5h10M5 2.5v9" style={{ ...muted, strokeWidth: 0.75 }} />
      <path d="M11.5 6.5l-2.5 4.25h2.25L10.25 15l4-5.25H12l1.25-3.25z" fill={accent} />
    </svg>
  )
}
