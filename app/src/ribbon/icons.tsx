import type { BorderPreset } from "@cells/core"
import type { ReactNode, SVGProps } from "react"

/**
 * Ribbon pictures drawn for Cells: one 16 px grid, 1.25 px strokes with round joins, colours from the icon tokens
 * (`--icon-base`, `--icon-accent`, `--icon-muted`) so they theme. lucide covers the common ones; these are the
 * spreadsheet-specific ones it lacks.
 */
function Svg({ children, size = 16, ...rest }: { children: ReactNode; size?: number } & SVGProps<SVGSVGElement>) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="var(--icon-base)"
      strokeWidth={1.25}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
      {...rest}
    >
      {children}
    </svg>
  )
}

const ACCENT = "var(--icon-accent)"
const MUTED = "var(--icon-muted)"

/** A cell square drawn dotted (the sides without a border), with the given sides solid. */
export function BorderIcon({ preset, size }: { preset: BorderPreset; size?: number }) {
  const dotted = { stroke: MUTED, strokeDasharray: "1 1.5", strokeLinecap: "butt" as const }
  const sides = {
    top: "M2.5 2.5h11",
    bottom: "M2.5 13.5h11",
    left: "M2.5 2.5v11",
    right: "M13.5 2.5v11",
    h: "M2.5 8h11",
    v: "M8 2.5v11",
  }
  const on: Record<BorderPreset, (keyof typeof sides)[]> = {
    bottom: ["bottom"],
    top: ["top"],
    left: ["left"],
    right: ["right"],
    none: [],
    all: ["top", "bottom", "left", "right", "h", "v"],
    outside: ["top", "bottom", "left", "right"],
    thickOutside: ["top", "bottom", "left", "right"],
    bottomDouble: ["bottom"],
    thickBottom: ["bottom"],
    topBottom: ["top", "bottom"],
    topThickBottom: ["top", "bottom"],
    topDoubleBottom: ["top", "bottom"],
    insideHorizontal: ["h"],
    insideVertical: ["v"],
  }
  const thick = preset === "thickOutside" ? on.thickOutside : preset.includes("hickBottom") ? ["bottom"] : []
  const double = preset === "bottomDouble" || preset === "topDoubleBottom"
  return (
    <Svg size={size}>
      {(Object.keys(sides) as (keyof typeof sides)[])
        .filter((s) => !on[preset].includes(s))
        .map((s) => (
          <path key={s} d={sides[s]} {...dotted} />
        ))}
      {on[preset].map((s) =>
        double && s === "bottom" ? (
          <path key={s} d="M2.5 12.5h11M2.5 14.5h11" strokeWidth={1} stroke="var(--icon-base)" />
        ) : (
          <path key={s} d={sides[s]} strokeWidth={thick.includes(s) ? 2.25 : 1.25} strokeLinecap="square" />
        ),
      )}
    </Svg>
  )
}

/** "A" over a colour bar (Font Color), the bar showing the colour the button applies. */
export function FontColorIcon({ color }: { color: string }) {
  return (
    <Svg>
      <path d="M4.5 10.5 8 2l3.5 8.5M5.7 7.6h4.6" />
      <rect x="2" y="12.5" width="12" height="2.5" rx="0.5" fill={color} stroke="none" />
    </Svg>
  )
}

/** A paint bucket over a colour bar (Fill Color). */
export function FillColorIcon({ color }: { color: string }) {
  return (
    <Svg>
      <path d="M3.5 6.5 7.5 2.5l4 4-4 4z" />
      <path
        d="M12.5 8.2c.6.8 1 1.4 1 1.9a1 1 0 0 1-2 0c0-.5.4-1.1 1-1.9z"
        fill={ACCENT}
        stroke={ACCENT}
        strokeWidth={0.8}
      />
      <rect x="2" y="12.5" width="12" height="2.5" rx="0.5" fill={color} stroke="none" />
    </Svg>
  )
}

export function IncreaseDecimalIcon() {
  return (
    <Svg>
      <path d="M2 9.5h.01" strokeWidth={1.8} />
      <text x="3.4" y="10" fontSize="5.6" fill="var(--icon-base)" stroke="none" fontFamily="sans-serif">
        .00
      </text>
      <path d="M3 13.5h7M8.5 12l1.5 1.5L8.5 15" stroke={ACCENT} />
      <text x="3" y="5.5" fontSize="5.6" fill="var(--icon-muted)" stroke="none" fontFamily="sans-serif">
        .0
      </text>
    </Svg>
  )
}

export function DecreaseDecimalIcon() {
  return (
    <Svg>
      <text x="2.6" y="5.5" fontSize="5.6" fill="var(--icon-muted)" stroke="none" fontFamily="sans-serif">
        .00
      </text>
      <text x="3.4" y="10" fontSize="5.6" fill="var(--icon-base)" stroke="none" fontFamily="sans-serif">
        .0
      </text>
      <path d="M10 13.5H3M4.5 12 3 13.5 4.5 15" stroke={ACCENT} />
    </Svg>
  )
}

export function CommaIcon() {
  return (
    <Svg>
      <text
        x="1.5"
        y="11.5"
        fontSize="10"
        fontWeight="600"
        fill="var(--icon-base)"
        stroke="none"
        fontFamily="sans-serif"
      >
        9,
      </text>
      <path d="M12.5 4v8" stroke={MUTED} />
    </Svg>
  )
}

export function AccountingIcon() {
  return (
    <Svg>
      <rect x="1.5" y="3.5" width="13" height="9" rx="1.5" />
      <path
        d="M9.3 6.1c-.4-.4-.9-.6-1.5-.6-.8 0-1.4.4-1.4 1s.5.9 1.4 1.1c.9.2 1.5.5 1.5 1.1s-.6 1-1.5 1c-.6 0-1.2-.2-1.6-.6M7.9 4.6v.9M7.9 10.7v.8"
        stroke={ACCENT}
      />
    </Svg>
  )
}

export function MergeIcon() {
  return (
    <Svg>
      <rect x="1.5" y="3.5" width="13" height="9" rx="0.5" stroke={MUTED} />
      <path d="M4 8h8M6 6 4 8l2 2M10 6l2 2-2 2" stroke={ACCENT} />
    </Svg>
  )
}

export function OrientationIcon() {
  return (
    <Svg>
      <path d="m3 13 5-9 2.5 4.5M4.8 9.8h4.4" />
      <path d="M11 13.5 14 10.5M14 13.5h-3v-3" stroke={ACCENT} />
    </Svg>
  )
}

export function ConditionalFormattingIcon() {
  return (
    <Svg size={32}>
      <rect x="1.5" y="1.5" width="13" height="13" rx="1" stroke={MUTED} />
      <path d="M1.5 6h13M1.5 10.5h13M6 1.5v13" stroke={MUTED} />
      <rect x="6.6" y="2.1" width="7.3" height="3.3" fill="#e05252" stroke="none" opacity={0.85} />
      <rect x="6.6" y="6.6" width="5" height="3.3" fill="#4f9fd8" stroke="none" opacity={0.85} />
      <rect x="6.6" y="11.1" width="2.6" height="3.3" fill="#56a96f" stroke="none" opacity={0.85} />
    </Svg>
  )
}

export function FormatAsTableIcon() {
  return (
    <Svg size={32}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <rect x="2" y="3" width="12" height="2.6" fill={ACCENT} stroke="none" opacity={0.85} />
      <path d="M1.5 8.2h13M1.5 10.9h13M6 5.6v8M10 5.6v8" stroke={MUTED} />
    </Svg>
  )
}

export function CellStylesIcon() {
  return (
    <Svg size={32}>
      <rect x="1.5" y="2.5" width="6" height="4.5" rx="0.6" fill="#c6efce" stroke="#4f9a5d" />
      <rect x="8.5" y="2.5" width="6" height="4.5" rx="0.6" fill="#ffc7ce" stroke="#b9545e" />
      <rect x="1.5" y="9" width="6" height="4.5" rx="0.6" fill="#ffeb9c" stroke="#a68a2b" />
      <rect x="8.5" y="9" width="6" height="4.5" rx="0.6" stroke={MUTED} />
    </Svg>
  )
}

/** Insert / Delete / Format in the Cells group: a sheet grid with a mark. */
export function CellsIcon({ kind }: { kind: "insert" | "delete" | "format" }) {
  return (
    <Svg size={32}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" stroke={MUTED} />
      <path d="M1.5 6.2h13M1.5 9.8h13M6 2.5v11M10 2.5v11" stroke={MUTED} />
      {kind === "insert" ? (
        <>
          <rect x="6" y="6.2" width="4" height="3.6" fill="var(--icon-accent-soft)" stroke={ACCENT} />
          <path d="M12.5 1v3M11 2.5h3" stroke={ACCENT} />
        </>
      ) : kind === "delete" ? (
        <>
          <rect x="6" y="6.2" width="4" height="3.6" fill="var(--icon-danger-soft)" stroke="var(--icon-danger)" />
          <path d="m11 1 3 3M14 1l-3 3" stroke="var(--icon-danger)" />
        </>
      ) : (
        <>
          <rect x="6" y="6.2" width="4" height="3.6" fill="var(--icon-accent-soft)" stroke={ACCENT} />
          <path d="M11.5 15.5v-2.2M11.5 13.3l2.5-2.5 1 1-2.5 2.5z" stroke="var(--icon-base)" strokeWidth={1} />
        </>
      )}
    </Svg>
  )
}

/** The picture a collapsed group shows when the window is too narrow for it. */
export function GroupIcon() {
  return (
    <Svg size={32}>
      <rect x="2.5" y="2.5" width="11" height="11" rx="1.5" stroke={MUTED} />
      <path d="M5 6h6M5 8h6M5 10h4" />
    </Svg>
  )
}

/** Placeholder picture for commands that are not built yet: a dotted square. */
export function PendingIcon({ size = 16 }: { size?: number }) {
  return (
    <Svg size={size}>
      <rect x="2.5" y="2.5" width="11" height="11" rx="2" stroke={MUTED} strokeDasharray="1.5 2" />
    </Svg>
  )
}
