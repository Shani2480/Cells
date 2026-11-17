import {
  SPARKLINE_STYLES,
  sparkColorHex,
  type SparkColors,
  type SparkAxisScale,
  type SparklineType,
  type Theme,
} from "@cells/core"
import type { SparkFlag } from "@/sheet/sparklines"
import { ColorPanel } from "./colors"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

const has = (env: RibbonEnv) => env.c.sparklines.active() !== null

/** A style's look in the gallery: a little line with its high and low points. */
export function SparkStyleThumb({ index, theme }: { index: number; theme: Theme }) {
  const c = SPARKLINE_STYLES[index]?.colors
  if (!c) return null
  const hex = (k: keyof typeof c) => sparkColorHex(c[k], theme)
  const pts = [10, 4, 13, 8, 2, 9]
  const x = (i: number) => 4 + i * 6.4
  const y = (v: number) => 3 + v * 1.2
  return (
    <svg width={40} height={24} viewBox="0 0 40 24" aria-hidden className="spark-thumb">
      <polyline
        points={pts.map((v, i) => `${x(i)},${y(v)}`).join(" ")}
        fill="none"
        stroke={hex("series")}
        strokeWidth={1.6}
        strokeLinejoin="round"
      />
      <circle cx={x(4)} cy={y(2)} r={2} fill={hex("high")} />
      <circle cx={x(2)} cy={y(13)} r={2} fill={hex("low")} />
    </svg>
  )
}

function StyleGallery({ env, close }: { env: RibbonEnv; close: () => void }) {
  const current = env.c.sparklines.styleIndex()
  return (
    <div className="spark-gallery" role="listbox" aria-label="Sparkline styles">
      {SPARKLINE_STYLES.map((s, i) => (
        <button
          key={s.name}
          type="button"
          role="option"
          aria-selected={i === current}
          className="spark-gallery-item"
          title={`Sparkline Style ${s.name}`}
          onClick={() => {
            env.c.sparklines.setStyle(i)
            close()
          }}
        >
          <SparkStyleThumb index={i} theme={env.c.wb.theme} />
        </button>
      ))}
    </div>
  )
}

function colorPanel(key: keyof SparkColors): MenuEntry {
  return {
    panel: (env, close) => {
      const now = env.c.sparklines.color(key)
      return (
        <ColorPanel
          theme={env.c.wb.theme}
          current={now ? sparkColorHex(now, env.c.wb.theme) : undefined}
          onPick={(hex, ref) => {
            if (hex) env.c.sparklines.setColor(key, ref ?? hex)
            close()
          }}
        />
      )
    },
  }
}

const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]

const TYPES: [SparklineType, string, string, string][] = [
  ["line", "SparklineTypeLine", "Line", "Change the sparkline group to a line sparkline."],
  ["column", "SparklineTypeColumn", "Column", "Change the sparkline group to a column sparkline."],
  ["stacked", "SparklineTypeWinLoss", "Win/Loss", "Change the sparkline group to a win/loss sparkline."],
]

const FLAGS: [SparkFlag, string, string, string, keyof SparkColors][] = [
  [
    "high",
    "SparklineShowHigh",
    "High Point",
    "Highlight the highest point(s) of the selected sparkline group.",
    "high",
  ],
  ["low", "SparklineShowLow", "Low Point", "Highlight the lowest point(s) of the selected sparkline group.", "low"],
  ["negative", "SparklineShowNegative", "Negative Points", "Show the negative points in another colour.", "negative"],
  ["first", "SparklineShowFirst", "First Point", "Highlight the first point of each sparkline.", "first"],
  ["last", "SparklineShowLast", "Last Point", "Highlight the last point of each sparkline.", "last"],
  ["markers", "SparklineShowMarkers", "Markers", "Show every point of the line sparklines.", "markers"],
]

const SCALES: [SparkAxisScale, string][] = [
  ["individual", "Automatic for Each Sparkline"],
  ["group", "Same for All Sparklines"],
  ["custom", "Custom Value..."],
]

/** Insert › Sparklines and the Sparkline tab (shown while the active cell has a sparkline). */
export function sparklineCommands(): Command[] {
  const insert = (id: string, type: SparklineType, description: string, keytip: string): Command => ({
    id,
    label: { line: "Line", column: "Column", stacked: "Win/Loss" }[type],
    description,
    keytip,
    run: (env) => env.c.sparklines.openCreate(type),
  })
  return [
    insert("SparklineLineInsert", "line", "Insert a line chart within a single cell.", "SL"),
    insert("SparklineColumnInsert", "column", "Insert a column chart within a single cell.", "SO"),
    insert("SparklineWinLossInsert", "stacked", "Insert a win/loss chart within a single cell.", "SW"),

    // ── Sparkline tab ──
    {
      id: "SparklineEditData",
      label: "Edit Data",
      description: "Edit the location and data source of the sparkline group, or of one sparkline.",
      keytip: "I",
      enabled: has,
      run: (env) => env.c.sparklines.openEditGroup(),
      menu: ["SparklineEditGroup", "SparklineEditSingle", "-", "SparklineHiddenEmpty"],
    },
    {
      id: "SparklineEditGroup",
      label: "Edit Group Location & Data...",
      enabled: has,
      run: (env) => env.c.sparklines.openEditGroup(),
    },
    {
      id: "SparklineEditSingle",
      label: "Edit Single Sparkline's Data...",
      enabled: has,
      run: (env) => env.c.sparklines.openEditSingle(),
    },
    {
      id: "SparklineHiddenEmpty",
      label: "Hidden & Empty Cells...",
      enabled: has,
      run: (env) => env.dialog("sparklineEmpty"),
    },
    ...TYPES.map(([type, id, label, description], i): Command => ({
      id,
      label,
      description,
      keytip: ["L", "O", "W"][i],
      enabled: has,
      pressed: (env) => env.c.sparklines.type() === type,
      run: (env) => env.c.sparklines.setType(type),
    })),
    ...FLAGS.map(([flag, id, label, description], i): Command => ({
      id,
      label,
      description,
      keytip: ["H", "V", "N", "F", "T", "M"][i],
      check: true,
      enabled: (env) => has(env) && (flag !== "markers" || env.c.sparklines.type() === "line"),
      pressed: (env) => env.c.sparklines.flag(flag),
      run: (env) => env.c.sparklines.toggleFlag(flag),
    })),
    {
      id: "SparklineStyleGallery",
      label: "Style",
      description: "Pick a visual style for the sparkline group.",
      keytip: "J",
      enabled: has,
      largeIcon: (env) => <SparkStyleThumb index={Math.max(0, env.c.sparklines.styleIndex())} theme={env.c.wb.theme} />,
      menu: [{ panel: (env, close) => <StyleGallery env={env} close={close} /> }],
    },
    {
      id: "SparklineColor",
      label: "Sparkline Color",
      description: "Pick the colour of the sparklines in the group, and the line's weight.",
      keytip: "R",
      enabled: has,
      menu: [colorPanel("series"), "-", "SparklineWeight"],
    },
    {
      id: "SparklineWeight",
      label: "Weight",
      enabled: (env) => env.c.sparklines.type() === "line",
      menu: WEIGHTS.map((w) => `SparklineWeight:${w}`),
    },
    ...WEIGHTS.map((w): Command => ({
      id: `SparklineWeight:${w}`,
      label: `${w} pt`,
      pressed: (env) => env.c.sparklines.weight() === w,
      run: (env) => env.c.sparklines.setWeight(w),
    })),
    {
      id: "SparklineMarkerColor",
      label: "Marker Color",
      description: "Pick the colours of the markers and the highlighted points.",
      keytip: "K",
      enabled: has,
      menu: FLAGS.map(([flag]) => `SparklineMarkerColor:${flag}`),
    },
    ...FLAGS.map(([flag, , label, , key]): Command => ({
      id: `SparklineMarkerColor:${flag}`,
      label,
      menu: [colorPanel(key)],
    })),
    {
      id: "SparklineAxis",
      label: "Axis",
      description: "The sparklines' axes: date axis, showing the horizontal axis, right to left, vertical bounds.",
      keytip: "A",
      enabled: has,
      menu: [
        { heading: "Horizontal Axis Options" },
        "SparklineAxisGeneral",
        "SparklineAxisDate",
        "SparklineAxisShow",
        "SparklineAxisRightToLeft",
        { heading: "Vertical Axis Minimum Value Options" },
        ...SCALES.map(([s]) => `SparklineAxisMin:${s}`),
        { heading: "Vertical Axis Maximum Value Options" },
        ...SCALES.map(([s]) => `SparklineAxisMax:${s}`),
      ],
    },
    {
      id: "SparklineAxisGeneral",
      label: "General Axis Type",
      pressed: (env) => !env.c.sparklines.active()?.dateRange,
      run: (env) => void env.c.sparklines.setDateRange(null),
    },
    {
      id: "SparklineAxisDate",
      label: "Date Axis Type...",
      pressed: (env) => Boolean(env.c.sparklines.active()?.dateRange),
      run: (env) => env.dialog("sparklineDates"),
    },
    {
      id: "SparklineAxisShow",
      label: "Show Axis",
      pressed: (env) => Boolean(env.c.sparklines.active()?.displayXAxis),
      run: (env) => env.c.sparklines.setAxisOption({ displayXAxis: !env.c.sparklines.active()?.displayXAxis }),
    },
    {
      id: "SparklineAxisRightToLeft",
      label: "Plot Data Right-to-Left",
      pressed: (env) => Boolean(env.c.sparklines.active()?.rightToLeft),
      run: (env) => env.c.sparklines.setAxisOption({ rightToLeft: !env.c.sparklines.active()?.rightToLeft }),
    },
    ...(["min", "max"] as const).flatMap((which) =>
      SCALES.map(([scale, label]): Command => ({
        id: `SparklineAxis${which === "min" ? "Min" : "Max"}:${scale}`,
        label,
        pressed: (env) => env.c.sparklines.active()?.[which === "min" ? "minAxis" : "maxAxis"] === scale,
        run: (env) =>
          scale === "custom" ? env.dialog(`sparklineAxis:${which}`) : env.c.sparklines.setScale(which, scale),
      })),
    ),
    {
      id: "SparklineGroup",
      label: "Group",
      description: "Group the selected sparklines so they share their formatting.",
      keytip: "G",
      enabled: (env) => env.c.sparklines.canGroup(),
      run: (env) => env.c.sparklines.group(),
    },
    {
      id: "SparklineUngroup",
      label: "Ungroup",
      description: "Ungroup the selected sparklines so each can be formatted on its own.",
      keytip: "U",
      enabled: (env) => env.c.sparklines.canUngroup(),
      run: (env) => env.c.sparklines.ungroup(),
    },
    {
      id: "SparklineClear",
      label: "Clear",
      description: "Clear the selected sparklines.",
      keytip: "C",
      enabled: has,
      run: (env) => env.c.sparklines.clear(),
      menu: ["SparklineClearSelected", "SparklineClearGroups"],
    },
    {
      id: "SparklineClearSelected",
      label: "Clear Selected Sparklines",
      enabled: has,
      run: (env) => env.c.sparklines.clear(),
    },
    {
      id: "SparklineClearGroups",
      label: "Clear Selected Sparkline Groups",
      enabled: has,
      run: (env) => env.c.sparklines.clear(true),
    },
  ]
}
