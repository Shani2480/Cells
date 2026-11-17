import { parseRange, type EditRange, type Range } from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import { at, attrText, kid, kids, type XNode } from "./dom"
import { readPasswordHash } from "./sheet-read"

/**
 * Review › Allow Edit Ranges in a worksheet: `<protectedRanges><protectedRange name sqref password|hash…/>`.
 * Excel 2010 and later may put the Windows permissions in a securityDescriptor child; they are kept as read.
 */
export function readEditRanges(node: XNode): EditRange[] {
  return kids(node, "protectedRange").flatMap((n) => {
    const name = at(n, "name")
    const ranges = (at(n, "sqref") ?? "").split(/\s+/).flatMap((ref) => parseRange(ref.replace(/\$/g, "")) ?? [])
    if (!name || !ranges.length) return []
    const e: EditRange = { name, ranges }
    const hash = readPasswordHash(n)
    if (hash) e.hash = hash
    else if (at(n, "password")) e.password = at(n, "password")!.toUpperCase()
    const security = at(n, "securityDescriptor") ?? kid(n, "securityDescriptor")?.text
    if (security) e.securityDescriptor = security
    return [e]
  })
}

export function editRangesXml(list: readonly EditRange[], rangeRef: (r: Range) => string): string {
  if (!list.length) return ""
  const items = list.map((e) => {
    const attrs = attrText({
      password: e.hash ? undefined : e.password,
      algorithmName: e.hash?.algorithm,
      hashValue: e.hash?.hash,
      saltValue: e.hash?.salt,
      spinCount: e.hash ? String(e.hash.spinCount) : undefined,
      sqref: e.ranges.map(rangeRef).join(" "),
      name: e.name,
    })
    return e.securityDescriptor
      ? `<protectedRange${attrs}><securityDescriptor>${escapeXmlText(e.securityDescriptor)}</securityDescriptor></protectedRange>`
      : `<protectedRange${attrs}/>`
  })
  return `<protectedRanges>${items.join("")}</protectedRanges>`
}
