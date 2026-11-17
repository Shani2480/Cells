import { parseRange, type Hyperlink } from "@cells/core"
import { at, attrText, kids, type XNode } from "./dom"
import { NS_R } from "./constants"

/**
 * A worksheet's `<hyperlinks>`: each `<hyperlink ref="A1" r:id="rId3" location="Sheet2!A1" tooltip="…"/>`. A link
 * out of the workbook keeps its address in the sheet's relationships (TargetMode="External"), so reading gives the
 * relationship id for the caller to look up.
 */
export type HyperlinkRead = { link: Hyperlink; relId?: string }

export function readHyperlinks(root: XNode): HyperlinkRead[] {
  const out: HyperlinkRead[] = []
  for (const node of kids(root, "hyperlink")) {
    const ref = at(node, "ref")
    // A ref may be a list ("A1 B3") in files from other programs; each part is its own link.
    for (const part of (ref ?? "").split(/\s+/)) {
      const range = part ? parseRange(part) : null
      if (!range) continue
      const link: Hyperlink = { range }
      const location = at(node, "location")
      const tooltip = at(node, "tooltip")
      const display = at(node, "display")
      if (location) link.location = location
      if (tooltip) link.tooltip = tooltip
      if (display) link.display = display
      const relId = at(node, "id", NS_R)
      out.push(relId ? { link, relId } : { link })
    }
  }
  return out
}

/** The `<hyperlinks>` element; `relIds[i]` is the relationship of link i's address (null when it has none). */
export function hyperlinksXml(
  links: readonly Hyperlink[],
  relIds: readonly (string | null)[],
  rangeRef: (r: Hyperlink["range"]) => string,
): string {
  if (!links.length) return ""
  const items = links.map((l, i) => {
    const rid = relIds[i]
    return `<hyperlink${attrText({ ref: rangeRef(l.range) })}${rid ? ` r:id="${rid}"` : ""}${attrText({
      location: l.location,
      tooltip: l.tooltip,
      display: l.display,
    })}/>`
  })
  return `<hyperlinks>${items.join("")}</hyperlinks>`
}

/**
 * The relationship target for an address, as Excel writes one: web and mail addresses as they are, Windows paths
 * as file:/// links, other paths (relative to the workbook) with forward slashes.
 */
export function hyperlinkTarget(address: string): string {
  if (/^[a-z][a-z0-9+.-]*:/i.test(address) && !/^[a-z]:[\\/]/i.test(address)) return address
  if (/^[a-z]:[\\/]/i.test(address)) return "file:///" + address
  if (address.startsWith("\\\\")) return "file:///" + address
  return address.replace(/\\/g, "/")
}

/** The address a relationship target gives back (file:/// links of Windows paths become the path again). */
export function hyperlinkAddress(target: string): string {
  const m = /^file:\/\/\/([a-z]:[\\/].*|\\\\.*)$/i.exec(target)
  if (m) return m[1]
  return target
}
