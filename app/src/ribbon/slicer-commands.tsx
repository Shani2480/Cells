import { SLICER_STYLES, slicerLook, type Theme } from "@cells/core"
import type { Command, RibbonEnv } from "./registry"

const has = (env: RibbonEnv) => env.c.slicers.selected().length > 0
const one = (env: RibbonEnv) => env.c.slicers.active() !== null

/** A style's look in the gallery: a little slicer with a header line and three buttons, two picked. */
export function SlicerStyleThumb({ style, theme }: { style: string; theme: Theme }) {
  const look = slicerLook(style, theme)
  const buttons = [look.selected, look.selected, look.unselected]
  return (
    <svg width={48} height={40} viewBox="0 0 48 40" aria-hidden className="slicer-thumb">
      <rect x={0.5} y={0.5} width={47} height={39} fill={look.background} stroke={look.frame} />
      <rect x={4} y={4} width={18} height={3} fill="#404040" />
      {look.divider ? <line x1={4} y1={9.5} x2={44} y2={9.5} stroke={look.divider} /> : null}
      {buttons.map((b, i) => (
        <rect key={i} x={4.5} y={12.5 + i * 9} width={39} height={7} rx={1.5} fill={b.fill} stroke={b.border} />
      ))}
    </svg>
  )
}

/** The Slicer Styles gallery: Light, Other and Dark, as Excel groups them. */
function StyleGallery({ env, close }: { env: RibbonEnv; close: () => void }) {
  const current = env.c.slicers.styleIndex()
  const label = (s: string) => s.replace(/^SlicerStyle(\w+?)(\d)$/, "Slicer Style $1 $2")
  return (
    <div className="slicer-gallery" role="listbox" aria-label="Slicer styles">
      {(["Light", "Other", "Dark"] as const).map((family) => (
        <div key={family} className="slicer-gallery-group">
          <div className="color-heading">{family}</div>
          <div className="slicer-gallery-row">
            {SLICER_STYLES.filter((s) => s.includes(family)).map((s) => (
              <button
                key={s}
                type="button"
                role="option"
                aria-selected={SLICER_STYLES.indexOf(s) === current}
                className="spark-gallery-item"
                title={label(s)}
                onClick={() => {
                  env.c.slicers.setStyle(s)
                  close()
                }}
              >
                <SlicerStyleThumb style={s} theme={env.c.wb.theme} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

/** Pixels as the tab shows them: inches (cm outside the US), as Shape Format's size boxes do. */
const metric = () => !/^en-US/.test(navigator.language ?? "en-US")
const showLength = (px: number) =>
  metric()
    ? `${(px / (96 / 2.54)).toFixed(2).replace(/\.?0+$/, "")} cm`
    : `${(px / 96).toFixed(2).replace(/\.?0+$/, "")}"`
function readLength(text: string): number | null {
  const n = parseFloat(text.replace(",", "."))
  if (!Number.isFinite(n) || n <= 0) return null
  const cm = /cm/.test(text) || (!/["in]/.test(text) && metric())
  return cm ? (n * 96) / 2.54 : n * 96
}

/** Insert › Slicer, Table Design › Insert Slicer, and the Slicer tab (while slicers are selected). */
export function slicerCommands(): Command[] {
  return [
    {
      id: "SlicerInsert",
      label: "Slicer",
      description: "Use a slicer to filter data visually. Slicers make it faster and easier to filter tables.",
      keytip: "SF",
      run: (env) => env.c.slicers.openInsert(),
    },
    {
      id: "TableInsertSlicer",
      label: "Insert Slicer",
      description: "Filter the table's data visually.",
      keytip: "I",
      enabled: (env) => env.c.tables.active() !== null,
      run: (env) => env.c.slicers.openInsert(),
    },

    // ── Slicer tab ──
    {
      id: "SlicerCaption",
      label: "Slicer Caption",
      description: "The text in the slicer's header.",
      keytip: "TC",
      enabled: one,
      field: {
        width: 110,
        value: (env) => env.c.slicers.active()?.caption ?? "",
        commit: (env, text) => void env.c.slicers.setCaption(text),
      },
    },
    {
      id: "SlicerReportConnections",
      label: "Report Connections",
      description: "Choose the PivotTables this slicer filters. A table's slicer filters only its table.",
      keytip: "RC",
      enabled: () => false,
    },
    {
      id: "SlicerSettings",
      label: "Slicer Settings",
      description: "The slicer's name, header, and how its items are sorted and shown.",
      keytip: "TS",
      enabled: one,
      run: (env) => env.dialog("slicerSettings"),
    },
    {
      id: "SlicerStyleGallery",
      label: "Slicer Styles",
      description: "Pick a style for the slicer.",
      keytip: "S",
      enabled: has,
      largeIcon: (env) => (
        <SlicerStyleThumb style={env.c.slicers.active()?.style ?? SLICER_STYLES[0]} theme={env.c.wb.theme} />
      ),
      menu: [{ panel: (env, close) => <StyleGallery env={env} close={close} /> }],
    },
    {
      id: "SlicerColumns",
      label: "Columns",
      description: "The number of columns of buttons.",
      keytip: "C",
      enabled: has,
      field: {
        width: 56,
        value: (env) => String(env.c.slicers.active()?.columns ?? ""),
        commit: (env, text) => void env.c.slicers.setColumns(Number(text)),
      },
    },
    {
      id: "SlicerButtonHeight",
      label: "Height",
      description: "The height of the slicer's buttons.",
      keytip: "BH",
      enabled: has,
      field: {
        width: 64,
        value: (env) => {
          const o = env.c.slicers.active()
          return o ? showLength(o.buttonHeight) : ""
        },
        commit: (env, text) => {
          const px = readLength(text)
          if (px) env.c.slicers.setButtonHeight(px)
        },
      },
    },
    {
      id: "SlicerButtonWidth",
      label: "Width",
      description: "The width of the slicer's buttons.",
      keytip: "BW",
      enabled: has,
      field: {
        width: 64,
        value: (env) => {
          const o = env.c.slicers.active()
          return o ? showLength(env.c.slicers.buttonWidth(o)) : ""
        },
        commit: (env, text) => {
          const px = readLength(text)
          if (px) env.c.slicers.setButtonWidth(px)
        },
      },
    },
  ]
}
