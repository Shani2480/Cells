import { useEffect, useRef } from "react"
import { drawIcon } from "@/grid/cf-icons"
import type { CellPainter } from "@/grid/draw-cells"
import { COLOR_SCALES, DATA_BAR_COLORS, ICON_SET_GROUPS } from "@/sheet/conditional"
import { ICON_SETS, type IconSetName } from "@cells/core"
import type { Command, RibbonEnv } from "./registry"

/** Opens a quick rule dialog ("cf:greaterThan", …). */
const quick = (id: string, label: string, kind: string, description?: string): Command => ({
  id,
  label,
  description,
  run: (env) => env.dialog(`cf:${kind}`),
})

/** Galleries preview a choice on the sheet while the pointer is over it (Excel's live preview). */
function previewHandlers(env: RibbonEnv, rule: () => Parameters<RibbonEnv["c"]["conditional"]["setPreview"]>[0]) {
  return {
    onMouseEnter: () => env.c.conditional.setPreview(rule()),
    onMouseLeave: () => env.c.conditional.setPreview(null),
    onFocus: () => env.c.conditional.setPreview(rule()),
    onBlur: () => env.c.conditional.setPreview(null),
  }
}

function DataBarSwatch({ color, gradient }: { color: string; gradient: boolean }) {
  const fill = gradient ? `linear-gradient(90deg, ${color}, color-mix(in srgb, ${color} 10%, white))` : color
  return (
    <span className="cf-bar-swatch" aria-hidden>
      {[0.9, 0.55, 0.75, 0.3].map((w, i) => (
        <span
          key={i}
          style={{ width: `${w * 100}%`, background: fill, borderColor: gradient ? color : "transparent" }}
        />
      ))}
    </span>
  )
}

function ScaleSwatch({ colors }: { colors: string[] }) {
  const cells = colors.length === 3 ? [colors[2], colors[1], colors[0]] : [colors[1], colors[0]]
  // Excel's swatch: a 3×3 grid, best values on top.
  return (
    <span className="cf-scale-swatch" aria-hidden>
      {[0, 1, 2].map((row) =>
        [0, 1, 2].map((col) => {
          const t = row / 2
          const color =
            cells.length === 3 ? cells[Math.round(t * 2)] : t < 0.5 ? cells[0] : t > 0.5 ? cells[1] : cells[0]
          return <span key={`${row}${col}`} style={{ background: color }} />
        }),
      )}
    </span>
  )
}

/** An icon set's icons, best first, drawn with the grid's own painter. */
function IconSetSwatch({ set }: { set: IconSetName }) {
  const ref = useRef<HTMLCanvasElement>(null)
  const count = ICON_SETS[set]
  useEffect(() => {
    const canvas = ref.current
    const ctx = canvas?.getContext("2d")
    if (!canvas || !ctx) return
    const dpr = window.devicePixelRatio || 1
    canvas.width = count * 18 * dpr
    canvas.height = 16 * dpr
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    for (let i = 0; i < count; i++) drawIcon(ctx as unknown as CellPainter, set, count - 1 - i, i * 18 + 1, 0, 16)
  }, [set, count])
  return <canvas ref={ref} className="cf-icon-swatch" style={{ width: count * 18, height: 16 }} aria-hidden />
}

export function conditionalCommands(): Command[] {
  return [
    {
      id: "ConditionalFormattingHighlightCellsMenu",
      label: "Highlight Cells Rules",
      menu: [
        "CFGreaterThan",
        "CFLessThan",
        "CFBetween",
        "CFEqualTo",
        "CFTextContains",
        "CFDateOccurring",
        "CFDuplicateValues",
        "-",
        "CFMoreHighlightRules",
      ],
    },
    quick("CFGreaterThan", "Greater Than...", "greaterThan"),
    quick("CFLessThan", "Less Than...", "lessThan"),
    quick("CFBetween", "Between...", "between"),
    quick("CFEqualTo", "Equal To...", "equal"),
    quick("CFTextContains", "Text that Contains...", "text"),
    quick("CFDateOccurring", "A Date Occurring...", "date"),
    quick("CFDuplicateValues", "Duplicate Values...", "duplicate"),
    { id: "CFMoreHighlightRules", label: "More Rules...", run: (env) => env.dialog("cfNew:contains") },
    {
      id: "ConditionalFormattingTopBottomMenu",
      label: "Top/Bottom Rules",
      menu: [
        "CFTop10",
        "CFTop10Percent",
        "CFBottom10",
        "CFBottom10Percent",
        "CFAboveAverage",
        "CFBelowAverage",
        "-",
        "CFMoreTopRules",
      ],
    },
    quick("CFTop10", "Top 10 Items...", "top"),
    quick("CFTop10Percent", "Top 10%...", "topPercent"),
    quick("CFBottom10", "Bottom 10 Items...", "bottom"),
    quick("CFBottom10Percent", "Bottom 10%...", "bottomPercent"),
    quick("CFAboveAverage", "Above Average...", "above"),
    quick("CFBelowAverage", "Below Average...", "below"),
    { id: "CFMoreTopRules", label: "More Rules...", run: (env) => env.dialog("cfNew:top") },
    {
      id: "ConditionalFormattingDataBarsGallery",
      label: "Data Bars",
      menu: [
        { heading: "Gradient Fill" },
        { panel: (env, close) => <BarGrid env={env} close={close} gradient /> },
        { heading: "Solid Fill" },
        { panel: (env, close) => <BarGrid env={env} close={close} gradient={false} /> },
      ],
    },
    {
      id: "ConditionalFormattingColorScalesGallery",
      label: "Color Scales",
      menu: [
        {
          panel: (env, close) => (
            <div className="cf-gallery is-scales" role="menu">
              {COLOR_SCALES.map((s) => (
                <button
                  key={s.name}
                  type="button"
                  role="menuitem"
                  className="cf-gallery-item"
                  title={s.name}
                  aria-label={s.name}
                  {...previewHandlers(env, () => scaleRule(s.colors))}
                  onClick={() => {
                    env.c.conditional.colorScale(s.colors)
                    close()
                  }}
                >
                  <ScaleSwatch colors={s.colors} />
                </button>
              ))}
            </div>
          ),
        },
      ],
    },
    {
      id: "ConditionalFormattingIconSetsGallery",
      label: "Icon Sets",
      menu: ICON_SET_GROUPS.flatMap((g) => [
        { heading: g.heading },
        {
          panel: (env: RibbonEnv, close: () => void) => (
            <div className="cf-gallery is-icons" role="menu">
              {g.sets.map((s) => (
                <button
                  key={s.set}
                  type="button"
                  role="menuitem"
                  className="cf-gallery-item"
                  title={s.name}
                  aria-label={s.name}
                  {...previewHandlers(env, () => iconRule(s.set))}
                  onClick={() => {
                    env.c.conditional.iconSet(s.set)
                    close()
                  }}
                >
                  <IconSetSwatch set={s.set} />
                </button>
              ))}
            </div>
          ),
        },
      ]),
    },
    {
      id: "ConditionalFormattingNewRule",
      label: "New Rule...",
      run: (env) => env.dialog("cfNew:contains"),
    },
    {
      id: "ConditionalFormattingClearMenu",
      label: "Clear Rules",
      menu: ["CFClearSelected", "CFClearSheet", "CFClearTable", "CFClearPivot"],
    },
    {
      id: "CFClearSelected",
      label: "Clear Rules from Selected Cells",
      run: (env) => env.c.conditional.clearSelected(),
    },
    { id: "CFClearSheet", label: "Clear Rules from Entire Sheet", run: (env) => env.c.conditional.clearSheet() },
    // Tables and PivotTables come in Phase 4.
    { id: "CFClearTable", label: "Clear Rules from This Table", enabled: () => false },
    { id: "CFClearPivot", label: "Clear Rules from This PivotTable", enabled: () => false },
    {
      id: "ConditionalFormattingsManage",
      label: "Manage Rules...",
      run: (env) => env.dialog("cfManage"),
    },
  ]
}

function BarGrid({ env, close, gradient }: { env: RibbonEnv; close: () => void; gradient: boolean }) {
  return (
    <div className="cf-gallery is-bars" role="menu">
      {DATA_BAR_COLORS.map((b) => {
        const name = `${b.name} Data Bar`
        return (
          <button
            key={b.name}
            type="button"
            role="menuitem"
            className="cf-gallery-item"
            title={name}
            aria-label={`${gradient ? "Gradient" : "Solid"} ${name}`}
            {...previewHandlers(env, () => ({
              kind: "dataBar",
              min: { type: "automatic" },
              max: { type: "automatic" },
              color: b.color,
              gradient,
            }))}
            onClick={() => {
              env.c.conditional.dataBar(b.color, gradient)
              close()
            }}
          >
            <DataBarSwatch color={b.color} gradient={gradient} />
          </button>
        )
      })}
    </div>
  )
}

function scaleRule(colors: string[]) {
  return {
    kind: "colorScale" as const,
    points:
      colors.length === 3
        ? [
            { at: { type: "min" as const }, color: colors[0] },
            { at: { type: "percentile" as const, value: "50" }, color: colors[1] },
            { at: { type: "max" as const }, color: colors[2] },
          ]
        : [
            { at: { type: "min" as const }, color: colors[0] },
            { at: { type: "max" as const }, color: colors[1] },
          ],
  }
}

function iconRule(set: IconSetName) {
  const count = ICON_SETS[set]
  return {
    kind: "iconSet" as const,
    set,
    reverse: false,
    thresholds: Array.from({ length: count - 1 }, (_, i) => ({
      type: "percent" as const,
      value: String(Math.round(((i + 1) * 100) / count)),
    })),
    showValue: true,
  }
}
