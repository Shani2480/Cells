import { CHART_FAMILIES, chartVariant, renderChart, type ChartSpec } from "@cells/chart"
import { memo, useMemo, useState } from "react"
import { ChartTypeIcon } from "@/ribbon/chart-commands"
import { CHART_ADVICE } from "@/sheet/chart-recommend"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** The families All Charts lists (Excel's, without Map: Cells doesn't draw map charts from a sheet). */
const FAMILIES = CHART_FAMILIES.filter((f) => f.id !== "map")

export type ChartDialogStart = { tab?: "recommended" | "all"; family?: string }

const PREVIEW = { w: 440, h: 280 }
const THUMB = { w: 150, h: 96 }

/**
 * Insert › Recommended Charts (and the Charts launcher, "More … Charts...") and Chart Design › Change Chart Type:
 * Excel's dialog. Recommended Charts lists the types that suit the data with a large preview and a line on what each
 * is for; All Charts has every type by family. The previews draw the selected data (or, changing a type, the chart).
 */
export function InsertChartDialog({
  controller,
  change = false,
  start = {},
  onClose,
}: {
  controller: SheetController
  change?: boolean
  start?: ChartDialogStart
  onClose: () => void
}) {
  const charts = controller.charts
  const recommended = useMemo(() => charts.recommended(change), [charts, change])
  const current = change ? (charts.variant() ?? "column") : null
  const familyOf = (id: string) => FAMILIES.find((f) => f.items.some((v) => v.id === id))?.id ?? "column"
  const [tab, setTab] = useState<"recommended" | "all">(start.family ? "all" : (start.tab ?? "recommended"))
  const [family, setFamily] = useState(start.family ?? familyOf(current ?? recommended[0] ?? "column"))
  const [picked, setPicked] = useState(
    start.family
      ? (FAMILIES.find((f) => f.id === start.family)?.items[0]?.id ?? "column")
      : (current ?? recommended[0]),
  )
  // A family asked for ("More Bar Charts...") opens on its first type; otherwise the chart's own type.
  const [pickedAll, setPickedAll] = useState(
    (start.family ? undefined : current) ?? FAMILIES.find((f) => f.id === family)?.items[0]?.id ?? "column",
  )
  const chosen = tab === "recommended" ? picked : pickedAll
  const label = chartVariant(chosen)?.label ?? chosen
  // Each type drawn once with the data (the thumbnails and the large preview share them).
  const previews = useMemo(() => new Map<string, ChartSpec>(), [])
  const previewOf = (id: string) => {
    let spec = previews.get(id)
    if (!spec) previews.set(id, (spec = charts.preview(id, change)))
    return spec
  }
  const accept = () => {
    if (change) charts.setVariant(chosen)
    else charts.insert(chosen)
    onClose()
  }
  const shownFamily = FAMILIES.find((f) => f.id === family) ?? FAMILIES[0]
  return (
    <Dialog title={change ? "Change Chart Type" : "Insert Chart"} width={760} onCancel={onClose} onAccept={accept}>
      <div className="fc-tabs" role="tablist">
        {(
          [
            ["recommended", "Recommended Charts"],
            ["all", "All Charts"],
          ] as const
        ).map(([id, text]) => (
          <button
            key={id}
            type="button"
            role="tab"
            aria-selected={tab === id}
            className="fc-tab"
            onClick={() => setTab(id)}
          >
            {text}
          </button>
        ))}
      </div>
      {tab === "recommended" ? (
        <div className="ic-page" role="tabpanel" aria-label="Recommended Charts">
          <div className="ic-thumbs" role="listbox" aria-label="Recommended Charts">
            {recommended.map((id) => (
              <button
                key={id}
                type="button"
                role="option"
                aria-selected={picked === id}
                aria-label={chartVariant(id)?.label ?? id}
                title={chartVariant(id)?.label ?? id}
                className="ic-thumb"
                onClick={() => setPicked(id)}
                onDoubleClick={() => {
                  setPicked(id)
                  if (change) charts.setVariant(id)
                  else charts.insert(id)
                  onClose()
                }}
              >
                <ChartPicture spec={previewOf(id)} {...THUMB} scale={2.2} />
              </button>
            ))}
          </div>
          <div className="ic-main">
            <div className="ic-name">{label}</div>
            <ChartPicture spec={previewOf(chosen)} {...PREVIEW} />
            <p className="ic-advice">{CHART_ADVICE[chosen] ?? ""}</p>
          </div>
        </div>
      ) : (
        <div className="ic-page" role="tabpanel" aria-label="All Charts">
          <nav className="ic-families" aria-label="Chart types">
            {FAMILIES.map((f) => (
              <button
                key={f.id}
                type="button"
                aria-pressed={family === f.id}
                className="ic-family"
                onClick={() => {
                  setFamily(f.id)
                  setPickedAll(f.items[0].id)
                }}
              >
                <ChartTypeIcon id={f.items[0].id} size={20} />
                <span>{f.label}</span>
              </button>
            ))}
          </nav>
          <div className="ic-main">
            <div className="ic-variants" role="listbox" aria-label={shownFamily.label}>
              {shownFamily.items.map((v) => (
                <button
                  key={v.id}
                  type="button"
                  role="option"
                  aria-selected={pickedAll === v.id}
                  aria-label={v.label}
                  title={v.label}
                  className="ic-variant"
                  onClick={() => setPickedAll(v.id)}
                  onDoubleClick={() => {
                    if (change) charts.setVariant(v.id)
                    else charts.insert(v.id)
                    onClose()
                  }}
                >
                  <ChartTypeIcon id={v.id} size={44} />
                </button>
              ))}
            </div>
            <div className="ic-name">{label}</div>
            <ChartPicture spec={previewOf(chosen)} {...PREVIEW} />
          </div>
        </div>
      )}
    </Dialog>
  )
}

/**
 * A chart drawn at a size (the renderer draws its inside; the frame is ours). `scale` draws it that many times
 * larger and shrinks it, so a thumbnail is the whole chart made small, as Excel's are, not a cramped one.
 */
const ChartPicture = memo(function ChartPicture({
  spec,
  w,
  h,
  scale = 1,
}: {
  spec: ChartSpec
  w: number
  h: number
  scale?: number
}) {
  const svg = useMemo(() => {
    try {
      return renderChart(spec, w * scale, h * scale)
    } catch {
      return ""
    }
  }, [spec, w, h, scale])
  return (
    <svg
      className="ic-picture"
      width={w}
      height={h}
      viewBox={`0 0 ${w * scale} ${h * scale}`}
      aria-hidden
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  )
})
