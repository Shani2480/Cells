import { TIMELINE_STYLES, timelineLook, type Theme } from "@cells/core"
import type { TimelineShowKey } from "@/sheet/timelines"
import type { Command, RibbonEnv } from "./registry"

const has = (env: RibbonEnv) => env.c.timelines.selected().length > 0
const one = (env: RibbonEnv) => env.c.timelines.active() !== null

/** A style's look in the gallery: a little timeline with a header line and a strip of bars, two picked. */
export function TimelineStyleThumb({ style, theme }: { style: string; theme: Theme }) {
  const look = timelineLook(style, theme)
  const id = `tlg-${style}`
  return (
    <svg width={48} height={40} viewBox="0 0 48 40" aria-hidden className="slicer-thumb">
      <defs>
        <linearGradient id={id} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor={look.picked.top} />
          <stop offset="1" stopColor={look.picked.bottom} />
        </linearGradient>
      </defs>
      <rect x={0.5} y={0.5} width={47} height={39} fill={look.background} stroke={look.frame} />
      <rect x={4} y={4} width={18} height={3} fill="#404040" />
      <line x1={4} y1={9.5} x2={44} y2={9.5} stroke={look.divider} />
      <rect x={4} y={14} width={14} height={2} fill={look.label} />
      {[0, 1, 2, 3].map((i) => (
        <rect
          key={i}
          x={4 + i * 10}
          y={22}
          width={9.5}
          height={8}
          fill={i === 1 || i === 2 ? `url(#${id})` : look.unpicked.fill}
        />
      ))}
    </svg>
  )
}

/** The Timeline Styles gallery: Light and Dark, as Excel groups them. */
function StyleGallery({ env, close }: { env: RibbonEnv; close: () => void }) {
  const current = env.c.timelines.styleIndex()
  const label = (s: string) => s.replace(/^TimeSlicerStyle(\w+?)(\d)$/, "Timeline Style $1 $2")
  return (
    <div className="slicer-gallery" role="listbox" aria-label="Timeline styles">
      {(["Light", "Dark"] as const).map((family) => (
        <div key={family} className="slicer-gallery-group">
          <div className="color-heading">{family}</div>
          <div className="slicer-gallery-row">
            {TIMELINE_STYLES.filter((s) => s.includes(family)).map((s) => (
              <button
                key={s}
                type="button"
                role="option"
                aria-selected={TIMELINE_STYLES.indexOf(s) === current}
                className="spark-gallery-item"
                title={label(s)}
                onClick={() => {
                  env.c.timelines.setStyle(s)
                  close()
                }}
              >
                <TimelineStyleThumb style={s} theme={env.c.wb.theme} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

const show = (id: string, key: TimelineShowKey, label: string, keytip: string, description: string): Command => ({
  id,
  label,
  description,
  keytip,
  check: true,
  enabled: has,
  run: (env) => void env.c.timelines.setShown(key, !env.c.timelines.shown(key)),
  pressed: (env) => env.c.timelines.shown(key),
})

/** Insert › Timeline, PivotTable Analyze › Insert Timeline, and the Timeline tab (while timelines are selected). */
export function timelineCommands(): Command[] {
  return [
    {
      id: "InsertTimeline",
      label: "Timeline",
      description:
        "Use a timeline to filter dates interactively. Timelines make it faster and easier to filter by dates.",
      keytip: "TI",
      run: (env) => env.c.timelines.openInsert(),
    },
    {
      id: "PivotInsertTimeline",
      label: "Insert Timeline",
      description: "Use a timeline to filter the PivotTable's dates by years, quarters, months or days.",
      keytip: "IT",
      enabled: (env) => env.c.pivots.active() !== null,
      run: (env) => env.c.timelines.openInsert(),
    },

    // ── Timeline tab ──
    {
      id: "TimelineCaption",
      label: "Timeline Caption",
      description: "The text in the timeline's header.",
      keytip: "TC",
      enabled: one,
      field: {
        width: 110,
        value: (env) => env.c.timelines.active()?.caption ?? "",
        commit: (env, text) => void env.c.timelines.setCaption(text),
      },
    },
    {
      id: "TimelineReportConnections",
      label: "Report Connections",
      description: "Choose the PivotTables this timeline filters.",
      keytip: "RC",
      enabled: one,
      run: (env) => env.dialog("timelineConnections"),
    },
    {
      id: "TimelineStyleGallery",
      label: "Timeline Styles",
      description: "Pick a style for the timeline.",
      keytip: "S",
      enabled: has,
      largeIcon: (env) => (
        <TimelineStyleThumb style={env.c.timelines.active()?.style ?? TIMELINE_STYLES[0]} theme={env.c.wb.theme} />
      ),
      menu: [{ panel: (env, close) => <StyleGallery env={env} close={close} /> }],
    },
    show("TimelineShowHeader", "header", "Header", "SH", "Show the timeline's caption and its Clear Filter button."),
    show("TimelineShowSelectionLabel", "selectionLabel", "Selection Label", "SL", "Show the dates picked."),
    show("TimelineShowScrollbar", "scrollbar", "Scrollbar", "SB", "Show the scroll bar under the periods."),
    show(
      "TimelineShowTimeLevel",
      "timeLevel",
      "Time Level",
      "TL",
      "Show the menu that picks years, quarters, months or days.",
    ),
  ]
}
