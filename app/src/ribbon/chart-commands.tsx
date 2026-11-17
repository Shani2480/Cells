import { themeAccents } from "@/sheet/charts"
import {
  palettesFor,
  CHART_STYLES,
  QUICK_LAYOUTS,
  chartElements,
  chartVariant,
  defaultChartSpec,
  renderChart,
  textAsBars,
  withStyle,
  type ChartSpec,
} from "@cells/chart"
import { memo, type ReactNode } from "react"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/** Excel's Insert galleries: a heading and the chart types under it. */
type Section = { heading: string; ids: string[] }

const GALLERIES: Record<string, { sections: Section[]; more: string }> = {
  ChartTypeColumnInsertGallery: {
    sections: [
      { heading: "2-D Column", ids: ["column", "column-stacked", "column-percent"] },
      {
        heading: "3-D Column",
        ids: ["column-3d-clustered", "column-3d-stacked", "column-3d-percent", "column-3d"],
      },
      { heading: "2-D Bar", ids: ["hbar", "hbar-stacked", "hbar-percent"] },
      { heading: "3-D Bar", ids: ["hbar-3d-clustered", "hbar-3d-stacked", "hbar-3d-percent"] },
    ],
    more: "ChartTypeColumnInsertDialog",
  },
  ChartTypeLineInsertGallery: {
    sections: [
      {
        heading: "2-D Line",
        ids: ["line", "line-stacked", "line-percent", "line-markers", "line-stacked-markers", "line-percent-markers"],
      },
      { heading: "3-D Line", ids: ["line-3d"] },
      { heading: "2-D Area", ids: ["area", "area-stacked", "area-percent"] },
      { heading: "3-D Area", ids: ["area-3d", "area-3d-stacked", "area-3d-percent"] },
    ],
    more: "ChartTypeLineInsertDialog",
  },
  ChartTypePieInsertGallery: {
    sections: [
      { heading: "2-D Pie", ids: ["pie", "pie-of-pie", "bar-of-pie"] },
      { heading: "3-D Pie", ids: ["pie-3d"] },
      { heading: "Doughnut", ids: ["doughnut"] },
    ],
    more: "ChartTypePieInsertDialog",
  },
  ChartTypeHierarchyInsertGallery: {
    sections: [
      { heading: "Treemap", ids: ["treemap"] },
      { heading: "Sunburst", ids: ["sunburst"] },
    ],
    more: "ChartTypeHierarchyInsertDialog",
  },
  ChartTypeStatisticalInsertGallery: {
    sections: [
      { heading: "Histogram", ids: ["histogram", "pareto"] },
      { heading: "Box and Whisker", ids: ["box-whisker"] },
    ],
    more: "ChartTypeStatisticalInsertDialog",
  },
  ChartTypeXYScatterInsertGallery: {
    sections: [
      {
        heading: "Scatter",
        ids: ["scatter", "scatter-smooth-markers", "scatter-smooth", "scatter-lines-markers", "scatter-lines"],
      },
    ],
    more: "ChartTypeScatterInsertDialog",
  },
  ChartTypeOtherInsertGallery: {
    sections: [
      { heading: "Waterfall", ids: ["waterfall"] },
      { heading: "Funnel", ids: ["funnel"] },
      { heading: "Stock", ids: ["stock-hlc", "stock-ohlc", "stock-vhlc", "stock-vohlc"] },
      {
        heading: "Surface",
        ids: ["surface-3d", "surface-wireframe", "surface-contour", "surface-wireframe-contour"],
      },
      { heading: "Radar", ids: ["radar", "radar-markers", "radar-filled"] },
    ],
    more: "ChartTypeStockInsertDialog",
  },
}

/** Excel's names for the galleries' last rows (the Insert Chart dialog they open isn't built yet). */
const MORE_LABELS: Record<string, string> = {
  ChartTypeColumnInsertDialog: "More Column Charts...",
  ChartTypeLineInsertDialog: "More Line Charts...",
  ChartTypePieInsertDialog: "More Pie Charts...",
  ChartTypeHierarchyInsertDialog: "More Hierarchy Charts...",
  ChartTypeStatisticalInsertDialog: "More Statistical Charts...",
  ChartTypeScatterInsertDialog: "More Scatter Charts...",
  ChartTypeStockInsertDialog: "More Stock Charts...",
}

/** Recommended Charts: the common types, as Excel offers them first for most data. */
const RECOMMENDED = [
  "column",
  "line",
  "pie",
  "hbar",
  "area",
  "scatter",
  "column-stacked",
  "doughnut",
  "combo-column-line",
]

/** A chart type's picture: the type drawn small, without its title, legend or axes, as Excel's gallery icons. */
export const ChartTypeIcon = memo(function ChartTypeIcon({ id, size = 40 }: { id: string; size?: number }) {
  const spec = defaultChartSpec(id)
  const plain: ChartSpec = {
    ...spec,
    title: { ...spec.title, show: false },
    legend: "none",
    dataLabels: "none",
    axes: { category: false, value: false },
    gridlines: { major: false, minor: false, vertical: false },
    axisTitles: {
      category: { ...spec.axisTitles.category, show: false },
      value: { ...spec.axisTitles.value, show: false },
    },
  }
  let svg: string
  try {
    svg = textAsBars(renderChart(plain, 160, 120))
  } catch {
    svg = ""
  }
  return (
    <svg
      className="chart-type-icon"
      width={size}
      height={(size * 3) / 4}
      viewBox="0 0 160 120"
      aria-hidden
      data-keep-light
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  )
})

/** A chart drawn small with its own look (Quick Layout, Chart Styles, Change Colors). */
export const ChartPreview = memo(function ChartPreview({
  spec,
  width = 96,
  height = 64,
}: {
  spec: ChartSpec
  width?: number
  height?: number
}) {
  const w = 320
  const h = Math.round((320 * height) / width)
  let svg: string
  try {
    svg = textAsBars(renderChart(spec, w, h))
  } catch {
    svg = ""
  }
  return (
    <svg
      className="chart-preview"
      width={width}
      height={height}
      viewBox={`0 0 ${w} ${h}`}
      aria-hidden
      data-keep-light
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  )
})

function galleryPanel(sections: Section[], pick: (env: RibbonEnv, id: string) => void): MenuEntry {
  return {
    panel: (env, close) => (
      <div className="chart-gallery" role="menu">
        {sections.map((s) => (
          <div key={s.heading}>
            <div className="color-heading">{s.heading}</div>
            <div className="chart-gallery-row">
              {s.ids.map((id) => {
                const label = chartVariant(id)?.label ?? id
                return (
                  <button
                    key={id}
                    type="button"
                    role="menuitem"
                    className="chart-gallery-item"
                    title={label}
                    aria-label={label}
                    onClick={() => {
                      close()
                      pick(env, id)
                    }}
                  >
                    <ChartTypeIcon id={id} />
                  </button>
                )
              })}
            </div>
          </div>
        ))}
      </div>
    ),
  }
}

const insert = (env: RibbonEnv, id: string) => void env.c.charts.insert(id)
const inChart = (env: RibbonEnv) => env.c.charts.active() !== null

/** A grid of chart previews to pick a change from (Quick Layout, Chart Styles, Change Colors). */
function previewGrid(
  items: { key: string; label: string; spec: (s: ChartSpec) => ChartSpec; swatch?: ReactNode }[],
  label: string,
  columns: number,
): MenuEntry {
  return {
    panel: (env, close) => {
      const spec = env.c.charts.spec()
      if (!spec) return null
      return (
        <div className="chart-gallery" role="menu" aria-label={label}>
          <div className="chart-preview-grid" style={{ gridTemplateColumns: `repeat(${columns}, auto)` }}>
            {items.map((item) => (
              <button
                key={item.key}
                type="button"
                role="menuitem"
                className="chart-gallery-item"
                title={item.label}
                aria-label={item.label}
                onClick={() => {
                  close()
                  env.c.charts.change(item.spec, label)
                }}
              >
                {item.swatch ?? <ChartPreview spec={item.spec(spec)} />}
              </button>
            ))}
          </div>
        </div>
      )
    },
  }
}

/** Insert › Charts and the Chart Design tab. */
export function chartCommands(): Command[] {
  const galleries: Command[] = Object.entries(GALLERIES).map(([id, g]) => ({
    id,
    label: id,
    menu: [galleryPanel(g.sections, insert), "-", g.more],
  }))
  // Labels and descriptions come from Excel's names (excel-labels.json) where the registry fills them in.
  const named: Record<string, { label: string; description: string; keytip: string }> = {
    ChartTypeColumnInsertGallery: {
      label: "Insert Column or Bar Chart",
      description: "Use this chart type to visually compare values across a few categories.",
      keytip: "C",
    },
    ChartTypeLineInsertGallery: {
      label: "Insert Line or Area Chart",
      description: "Use this chart type to show trends over time (years, months, and days) or categories.",
      keytip: "N",
    },
    ChartTypePieInsertGallery: {
      label: "Insert Pie or Doughnut Chart",
      description: "Use this chart type to show proportions of a whole.",
      keytip: "Q",
    },
    ChartTypeHierarchyInsertGallery: {
      label: "Insert Hierarchy Chart",
      description: "Use this chart type to visually compare parts of a whole.",
      keytip: "I",
    },
    ChartTypeStatisticalInsertGallery: {
      label: "Insert Statistic Chart",
      description: "Use this chart type to show a statistical analysis of your data.",
      keytip: "H",
    },
    ChartTypeXYScatterInsertGallery: {
      label: "Insert Scatter (X, Y) or Bubble Chart",
      description: "Use this chart type to show the relationship between sets of values.",
      keytip: "D",
    },
    ChartTypeOtherInsertGallery: {
      label: "Insert Waterfall, Funnel, Stock, Surface, or Radar Chart",
      description: "Use this chart type to show the values of a financial or scientific data set.",
      keytip: "I",
    },
  }
  for (const g of galleries) Object.assign(g, named[g.id])
  // "More … Charts...": the Insert Chart dialog isn't built; Recommended Charts offers every type meanwhile.
  const more = Object.values(GALLERIES).map((g): Command => ({ id: g.more, label: MORE_LABELS[g.more] ?? g.more }))
  return [
    ...galleries,
    ...more,
    {
      id: "ChartInsertGalleryNew",
      label: "Recommended Charts",
      description: "Want us to recommend a good chart to showcase your data? Just select data in your sheet.",
      keytip: "R",
      menu: [galleryPanel([{ heading: "Recommended Charts", ids: RECOMMENDED }], insert)],
    },

    // ── Chart Design ──
    {
      id: "ChartAddElementMenu",
      label: "Add Chart Element",
      description: "Add elements to the chart, such as titles, legends, gridlines and data labels.",
      keytip: "A",
      enabled: inChart,
      menu: [
        {
          panel: (env, close) => {
            const spec = env.c.charts.spec()
            if (!spec) return null
            return (
              <div className="chart-elements" role="menu">
                {chartElements(spec)
                  .filter((e) => !e.disabled)
                  .map((e) => (
                    <div key={e.id} className="chart-element">
                      <div className="color-heading">{e.label}</div>
                      <div className="chart-element-choices">
                        {e.choices.map((choice) => (
                          <button
                            key={choice.label}
                            type="button"
                            role="menuitemcheckbox"
                            aria-checked={choice.on(spec)}
                            className={`chart-element-choice${choice.on(spec) ? " is-on" : ""}`}
                            onClick={() => {
                              close()
                              env.c.charts.change(choice.change, e.label)
                            }}
                          >
                            {choice.label}
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
              </div>
            )
          },
        },
      ],
    },
    {
      id: "ChartQuickLayoutGallery",
      label: "Quick Layout",
      description: "Change the overall layout of the chart.",
      keytip: "L",
      enabled: inChart,
      menu: [
        previewGrid(
          QUICK_LAYOUTS.map((l) => ({ key: String(l.id), label: l.label, spec: l.apply })),
          "Quick Layout",
          4,
        ),
      ],
    },
    {
      id: "ChartColorsGallery",
      label: "Change Colors",
      description: "Choose a different color set for the chart.",
      keytip: "H",
      enabled: inChart,
      menu: [
        {
          panel: (env, close) => {
            const spec = env.c.charts.spec()
            if (!spec) return null
            return (
              <div className="chart-gallery chart-colors" role="menu">
                {(["Colorful", "Monochromatic"] as const).map((group) => (
                  <div key={group}>
                    <div className="color-heading">{group}</div>
                    {palettesFor({ accents: themeAccents(env.c.wb.theme) })
                      .filter((p) => p.group === group)
                      .map((p) => (
                        <button
                          key={p.id}
                          type="button"
                          role="menuitemradio"
                          aria-checked={spec.palette === p.id}
                          className={`chart-palette${spec.palette === p.id ? " is-on" : ""}`}
                          title={p.label}
                          onClick={() => {
                            close()
                            env.c.charts.change((s) => ({ ...s, palette: p.id }), "Change Colors")
                          }}
                        >
                          {p.colors.slice(0, 6).map((c, i) => (
                            <span key={i} style={{ background: c }} />
                          ))}
                        </button>
                      ))}
                  </div>
                ))}
              </div>
            )
          },
        },
      ],
    },
    {
      id: "ChartStylesGallery",
      label: "Chart Styles",
      description: "Change the overall visual style of the chart.",
      keytip: "S",
      enabled: inChart,
      menu: [
        previewGrid(
          CHART_STYLES.map((_, i) => ({
            key: String(i + 1),
            label: `Style ${i + 1}`,
            spec: (s: ChartSpec) => withStyle(s, i + 1),
          })),
          "Chart Styles",
          4,
        ),
      ],
    },
    {
      id: "ChartSwitchRowColumn",
      label: "Switch Row/Column",
      description: "Swap the data over the axis: data in rows moves to columns, and the other way round.",
      keytip: "W",
      enabled: (env) => Boolean(env.c.charts.active()?.refs),
      run: (env) => env.c.charts.switchRowColumn(),
    },
    {
      id: "ChartSelectData",
      label: "Select Data",
      description: "Change the data range included in the chart.",
      keytip: "E",
      enabled: (env) => inChart(env) && Boolean(env.c.shell.chartData),
      run: (env) => env.c.shell.chartData?.(),
    },
    {
      id: "ChartChangeType",
      label: "Change Chart Type",
      description: "Change to a different type of chart.",
      keytip: "C",
      enabled: inChart,
      menu: [
        galleryPanel(
          Object.values(GALLERIES).flatMap((g) => g.sections),
          (env, id) => env.c.charts.setVariant(id),
        ),
      ],
    },
    {
      id: "ChartMove",
      label: "Move Chart",
      description: "Move this chart to another sheet in the workbook.",
      keytip: "M",
    },
  ]
}
