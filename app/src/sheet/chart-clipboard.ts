/**
 * A copied chart's HTML. Inkspire reads it as its own chart (the figure markup its editor writes: `data-figure`,
 * `data-kind="chart"` and the whole spec in `data-chart`), so the chart pastes there editable. The picture inside is
 * what Word, LibreOffice and mail paste. `data-cells-chart` says where it came from, for Paste Link later and for a
 * paste back into Cells (docs/inkspire-clipboard.md).
 */
import type { ChartSpec } from "@cells/chart"

export type CopiedChartInfo = {
  app: "cells"
  version: 1
  /** Which copy this is, so Cells knows its own chart when it comes back. */
  token: string
  file: string | null
  sheet: string
  name: string
  /** The chart's data range ("=Sheet1!$A$1:$C$4"), "" when it has none. */
  range: string
}

const escapeAttr = (text: string) =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

/** Inkspire draws charts at 96 px to the inch, as Cells does; sizes go over as they are. */
export function chartHtml(
  spec: ChartSpec,
  variant: string,
  size: { w: number; h: number },
  picture: string,
  info: CopiedChartInfo,
): string {
  const w = Math.round(size.w)
  const h = Math.round(size.h)
  const title = spec.title?.show && spec.title.text ? spec.title.text : info.name
  return (
    `<meta charset="utf-8"><div data-figure="" data-kind="chart" data-variant="${escapeAttr(variant)}"` +
    ` data-chart="${escapeAttr(JSON.stringify(spec))}" data-width="${w}" data-height="${h}"` +
    ` data-alt="${escapeAttr(title)}" data-cells-chart="${escapeAttr(JSON.stringify(info))}">` +
    `<img src="${escapeAttr(picture)}" width="${w}" height="${h}" alt="${escapeAttr(title)}"></div>`
  )
}

/** The `data-cells-chart` note of pasted HTML, when it's a chart Cells copied. */
export function copiedChartInfo(html: string): CopiedChartInfo | null {
  const m = /data-cells-chart="([^"]*)"/.exec(html)
  if (!m) return null
  try {
    const json = m[1]
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
    const info = JSON.parse(json) as CopiedChartInfo
    return info?.app === "cells" ? info : null
  } catch {
    return null
  }
}

/** A chart as an SVG picture (the fallback when no PNG of it is ready). */
export function svgPicture(svg: string, size: { w: number; h: number }): string {
  const markup = `<svg xmlns="http://www.w3.org/2000/svg" width="${size.w}" height="${size.h}" viewBox="0 0 ${size.w} ${size.h}">${svg}</svg>`
  return "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(markup)))
}
